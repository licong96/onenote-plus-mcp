import { getPageContent, listAllPagesInSection } from './pages.js';
import { getSection } from './sections.js';
import { collectNotebookTree, type SectionNode } from './tree.js';
import { searchIndex, readIndex, type IndexedPage } from '@/index-store/store.js';
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
 * Two candidate sources, tried in order:
 *   1. the local index (~/.config/onenote-mcp/index.json) — instant, offline;
 *   2. a live per-section scan — correct but expensive (one Graph round trip
 *      per section, each measured at 2–7s, so an account-wide scan runs into
 *      minutes and blows past the MCP request timeout).
 *
 * The index is therefore the default path and the live scan is the fallback.
 * Titles come free with page metadata; content requires fetching each page, so
 * it is opt-in and bounded.
 */

export interface FindPagesOptions {
  query: string;
  sectionId?: string;
  notebookId?: string;
  limit?: number;
  includeContent?: boolean;
  maxContentPages?: number;
  concurrency?: number;
  /** Set false to always hit Graph instead of the local index. */
  useIndex?: boolean;
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
  /** Where the candidate pages came from. */
  source?: 'index' | 'graph';
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

interface IndexSearchContext {
  needle: string;
  rawQuery: string;
  options: FindPagesOptions;
  limit: number;
  includeContent: boolean;
  maxContentPages: number;
  concurrency: number;
  scope: ScopedSection[];
  notes: string[];
}

const inScope = (page: IndexedPage, scope: ScopedSection[]): boolean =>
  scope.some((entry) => entry.section.id === page.sectionId);

/**
 * Derive the search scope from the index alone, with no network calls.
 *
 * Returns undefined when the caller asked for a notebook or section the index
 * does not know about — the signal to fall back to a live scan rather than
 * answer "no matches" from an incomplete mirror.
 */
const scopeFromIndexArgs = (
  index: Parameters<typeof searchIndex>[0],
  options: FindPagesOptions,
): ScopedSection[] | undefined => {
  if (options.sectionId) {
    const owner = index.notebooks.find((notebook) =>
      notebook.sections.some((section) => section.id === options.sectionId),
    );
    const section = owner?.sections.find((candidate) => candidate.id === options.sectionId);
    if (!owner || !section) return undefined;
    return [
      {
        notebook: owner.name,
        section: {
          id: section.id,
          name: section.name,
          isDefault: false,
          lastModified: section.lastModifiedDateTime,
          groupPath: section.groupPath,
        },
      },
    ];
  }

  const notebooks = options.notebookId
    ? index.notebooks.filter((notebook) => notebook.id === options.notebookId)
    : index.notebooks;
  if (notebooks.length === 0) return undefined;

  return notebooks.flatMap((notebook) =>
    notebook.sections.map((section) => ({
      notebook: notebook.name,
      section: {
        id: section.id,
        name: section.name,
        isDefault: false,
        lastModified: section.lastModifiedDateTime,
        groupPath: section.groupPath,
      },
    })),
  );
};

/**
 * Serve a search entirely from the local index.
 *
 * Scope filtering happens against page.sectionId rather than by resolving
 * notebook/section names, because scope resolution is itself a Graph call and
 * the index already carries every section it knows about.
 */
const searchFromIndex = async (
  index: Parameters<typeof searchIndex>[0],
  context: IndexSearchContext,
): Promise<FindPagesResult> => {
  const { needle, options, limit, includeContent, maxContentPages, concurrency, scope, notes } =
    context;

  const scoped = scope.filter((entry) => entry.section.id.length > 0);
  const candidates = index.pages.filter((page) => inScope(page, scoped));

  const titleMatches = candidates.filter((page) => page.title.toLowerCase().includes(needle));
  const contentCandidates = includeContent
    ? candidates.filter((page) => !page.title.toLowerCase().includes(needle))
    : [];

  const matches: PageMatch[] = titleMatches.map((page) => ({
    id: page.id,
    title: page.title,
    notebook: page.notebookName,
    section: page.sectionName,
    groupPath: page.groupPath.length > 0 ? page.groupPath : undefined,
    lastModified: page.lastModifiedDateTime,
    webUrl: page.webUrl,
    matchedIn: 'title',
  }));

  let contentsFetched = 0;
  const contentScanTruncated = contentCandidates.length > maxContentPages;
  if (includeContent) {
    // Still needs Graph: bodies are not mirrored. Bounded so a broad query
    // cannot turn into thousands of downloads.
    const budget = contentCandidates.slice(0, maxContentPages);
    const results = await mapWithConcurrency(budget, concurrency, async (page) => {
      try {
        const text = htmlToText(await getPageContent(page.id));
        contentsFetched += 1;
        if (!text.toLowerCase().includes(needle)) return null;
        return {
          id: page.id,
          title: page.title,
          notebook: page.notebookName,
          section: page.sectionName,
          groupPath: page.groupPath.length > 0 ? page.groupPath : undefined,
          lastModified: page.lastModifiedDateTime,
          webUrl: page.webUrl,
          matchedIn: 'content' as const,
          snippet: snippetAround(text, needle),
        };
      } catch {
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
  }

  // Newest-first, matching the live-scan path's ordering.
  matches.sort((left, right) =>
    (right.lastModified ?? '').localeCompare(left.lastModified ?? ''),
  );

  return {
    matches: matches.slice(0, limit),
    source: 'index',
    scanned: {
      notebooks: new Set(scoped.map((entry) => entry.notebook ?? entry.section.id)).size,
      sections: scoped.length,
      pages: candidates.length,
      contentsFetched,
    },
    contentScanTruncated,
    notes,
  };
};

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

  // Preferred path: answer from the local index. Checked BEFORE resolving scope,
  // because scope resolution itself walks Graph (172 section requests without a
  // notebook/section hint) — doing it first would pay the full cost the index
  // exists to avoid, and would trip the rate limiter even on a cache hit.
  if (options.useIndex !== false) {
    const index = await readIndex();
    if (index) {
      const scopeFromIndex = scopeFromIndexArgs(index, options);
      if (scopeFromIndex) {
        return searchFromIndex(index, {
          needle,
          rawQuery: options.query,
          options,
          limit,
          includeContent,
          maxContentPages,
          concurrency,
          scope: scopeFromIndex,
          notes,
        });
      }
      notes.push(
        'The given notebook/section is not in the local index; fell back to Graph. Run index mode "sync" or "rebuild" to refresh it.',
      );
    } else {
      notes.push(
        'No local index; fell back to scanning Graph (slow). Run the `index` tool with mode "rebuild" to make this instant.',
      );
    }
  }

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
    source: 'graph',
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
