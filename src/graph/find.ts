import { getPageContent, listAllPagesInSection } from './pages.js';
import { getSection } from './sections.js';
import { collectNotebookTree, type SectionNode } from './tree.js';
import { DEFAULT_CONCURRENCY, mapWithConcurrency } from '@/util/concurrency.js';
import type { Page } from './types.js';

/**
 * Scoped page search.
 *
 * Why this exists: Graph's own full-text search (`$search` on the pages
 * collection) is unusable from this account — the endpoint answers
 * "Your request contains unsupported OData query parameters" for `$search`,
 * section-scoped or not, and the account-wide `/me/onenote/pages` collection is
 * rejected outright with "The number of maximum sections is exceeded"
 * (this account has 172 sections). So `search_pages` — the upstream tool that
 * relies on both — cannot work here.
 *
 * Instead of a broken server-side index, this walks the sections in scope and
 * matches client-side. Titles come free with the page metadata; content
 * requires fetching each page, so it is opt-in and bounded.
 */

export interface FindPagesOptions {
  query: string;
  sectionId?: string;
  notebookId?: string;
  limit?: number;
  includeContent?: boolean;
  maxContentPages?: number;
  concurrency?: number;
}

export interface PageMatch {
  id: string;
  title: string;
  notebook?: string;
  section?: string;
  groupPath?: string[];
  lastModified?: string;
  webUrl?: string;
  matchedIn: 'title' | 'content';
  snippet?: string;
}

export interface FindPagesResult {
  matches: PageMatch[];
  scanned: {
    notebooks: number;
    sections: number;
    pages: number;
    contentsFetched: number;
  };
  /** True when the content pass stopped early at `maxContentPages`. */
  contentScanTruncated: boolean;
  notes: string[];
}

const CONTENT_PAGE_CEILING = 300;

const htmlToText = (html: string): string =>
  html
    .replaceAll(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, ' ')
    .replaceAll(/<br\s*\/?>/gi, ' ')
    .replaceAll(/<[^>]+>/g, ' ')
    .replaceAll('&nbsp;', ' ')
    .replaceAll('&amp;', '&')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&#39;', "'")
    .replaceAll(/\s+/g, ' ')
    .trim();

const snippetAround = (text: string, needle: string, radius = 70): string => {
  const index = text.toLowerCase().indexOf(needle);
  if (index < 0) return text.slice(0, radius * 2);
  const start = Math.max(0, index - radius);
  const end = Math.min(text.length, index + needle.length + radius);
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`;
};

interface ScopedSection {
  notebook?: string;
  section: SectionNode;
}

const resolveScope = async (
  options: FindPagesOptions,
  notes: string[],
): Promise<ScopedSection[]> => {
  if (options.sectionId) {
    const section = await getSection(options.sectionId);
    return [
      {
        notebook: section.parentNotebook?.displayName,
        section: {
          id: section.id,
          name: section.displayName,
          isDefault: section.isDefault,
          lastModified: section.lastModifiedDateTime,
          groupPath: section.parentSectionGroup ? [section.parentSectionGroup.displayName] : [],
        },
      },
    ];
  }

  const tree = await collectNotebookTree(options.notebookId, options.concurrency ?? DEFAULT_CONCURRENCY);
  if (!options.notebookId) {
    notes.push(
      'No notebookId or sectionId given, so the entire account was scanned (13 notebooks / 172 sections). This is slow — scope it down when you can.',
    );
  }
  return tree.flatMap((notebook) =>
    notebook.sections.map((section) => ({ notebook: notebook.name, section })),
  );
};

export const findPages = async (options: FindPagesOptions): Promise<FindPagesResult> => {
  const notes: string[] = [];
  const needle = options.query.trim().toLowerCase();
  if (needle.length === 0) {
    throw new Error('query must be at least one character.');
  }

  const limit = Math.max(1, options.limit ?? 25);
  const includeContent = options.includeContent ?? false;
  const maxContentPages = Math.min(
    Math.max(1, options.maxContentPages ?? 100),
    CONTENT_PAGE_CEILING,
  );
  const concurrency = options.concurrency ?? DEFAULT_CONCURRENCY;

  const scope = await resolveScope(options, notes);

  // Pass 1 — titles. Metadata only, one request per section.
  const scannedPerSection = await mapWithConcurrency(
    scope,
    concurrency,
    async ({ notebook, section }): Promise<{ pages: Page[]; notebook?: string; section: SectionNode }> => ({
      pages: await listAllPagesInSection(section.id),
      notebook,
      section,
    }),
  );

  const totalPages = scannedPerSection.reduce((sum, entry) => sum + entry.pages.length, 0);
  const matches: PageMatch[] = [];
  const contentCandidates: Array<{ page: Page; notebook?: string; section: SectionNode }> = [];

  for (const { pages, notebook, section } of scannedPerSection) {
    for (const page of pages) {
      const base: PageMatch = {
        id: page.id,
        title: page.title,
        notebook,
        section: section.name,
        groupPath: section.groupPath.length > 0 ? section.groupPath : undefined,
        lastModified: page.lastModifiedDateTime,
        webUrl: page.links?.oneNoteWebUrl?.href,
        matchedIn: 'title',
      };
      if (page.title?.toLowerCase().includes(needle)) {
        matches.push(base);
      } else if (includeContent) {
        contentCandidates.push({ page, notebook, section });
      }
    }
  }

  // Pass 2 — content, opt-in and bounded, so a broad query cannot turn into
  // thousands of page downloads.
  let contentsFetched = 0;
  const contentScanTruncated = contentCandidates.length > maxContentPages;
  if (includeContent) {
    const budget = contentCandidates.slice(0, maxContentPages);
    const results = await mapWithConcurrency(budget, concurrency, async (candidate) => {
      try {
        const text = htmlToText(await getPageContent(candidate.page.id));
        contentsFetched += 1;
        if (!text.toLowerCase().includes(needle)) return null;
        return {
          id: candidate.page.id,
          title: candidate.page.title,
          notebook: candidate.notebook,
          section: candidate.section.name,
          groupPath: candidate.section.groupPath.length > 0 ? candidate.section.groupPath : undefined,
          lastModified: candidate.page.lastModifiedDateTime,
          webUrl: candidate.page.links?.oneNoteWebUrl?.href,
          matchedIn: 'content' as const,
          snippet: snippetAround(text, needle),
        };
      } catch {
        // A single unreadable page shouldn't sink the whole search.
        return null;
      }
    });
    for (const result of results) {
      if (result) matches.push(result);
    }
    if (contentScanTruncated) {
      notes.push(
        `Content search stopped at ${maxContentPages} pages (${contentCandidates.length} candidates). Narrow the scope or raise maxContentPages.`,
      );
    }
  } else if (contentCandidates.length === 0 && matches.length === 0) {
    notes.push('Title search found nothing. Pass includeContent: true to also search page bodies.');
  }

  matches.sort((a, b) => {
    const left = a.lastModified ?? '';
    const right = b.lastModified ?? '';
    return right.localeCompare(left);
  });

  return {
    matches: matches.slice(0, limit),
    scanned: {
      notebooks: new Set(scope.map((entry) => entry.notebook ?? entry.section.id)).size,
      sections: scope.length,
      pages: totalPages,
      contentsFetched,
    },
    contentScanTruncated,
    notes,
  };
};
