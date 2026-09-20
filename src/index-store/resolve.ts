import { getPage } from '@/graph/pages.js';
import { listAllPagesInSection } from '@/graph/pages.js';
import {
  findPageByPath,
  indexStats,
  readIndex,
  searchIndex,
  writeIndex,
  type IndexedPage,
  type PageIndex,
} from './store.js';
import { syncIndex } from './sync.js';

/**
 * Resolve a human description of a page into a concrete page, repairing the
 * index along the way.
 *
 * The index is a cache, not a source of truth: page IDs change when a page is
 * moved or copied, so a stored ID can 404. Every resolution therefore falls
 * back to a live per-section lookup, and any repair is written back so the next
 * lookup is cheap again.
 */

export interface ResolveRequest {
  pageId?: string;
  /** Page title as shown in OneNote. */
  title?: string;
  /** Notebook name — scopes the lookup and doubles as a repair hint. */
  notebook?: string;
  /** Section name — scopes the lookup and doubles as a repair hint. */
  section?: string;
}

export interface ResolveResult {
  page?: IndexedPage;
  /** How the page was found, so callers can tell a cache hit from a repair. */
  source: 'index' | 'graph' | 'none';
  /** True when the resolution rewrote the on-disk index. */
  repaired: boolean;
  candidates: IndexedPage[];
  notes: string[];
  stats: ReturnType<typeof indexStats>;
}

const emptyStats = () => ({ notebooks: 0, sections: 0, pages: 0 });

export const resolvePage = async (
  request: ResolveRequest,
  options: { indexPath?: string } = {},
): Promise<ResolveResult> => {
  const notes: string[] = [];
  const index = (await readIndex(options.indexPath)) ?? undefined;
  if (!index) {
    return {
      source: 'none',
      repaired: false,
      candidates: [],
      notes: ['No index on disk yet. Run index_rebuild once so lookups can be served locally.'],
      stats: emptyStats(),
    };
  }

  const stats = indexStats(index);

  // 1. Exact id — O(1) when the index is current.
  if (request.pageId) {
    const hit = index.pages.find((page) => page.id === request.pageId);
    if (hit) return { page: hit, source: 'index', repaired: false, candidates: [hit], notes, stats };
    // The id is unknown (moved page, or index predates it) — verify live, since
    // a 404 here is meaningful rather than a cache miss.
    notes.push(`Page id ${request.pageId} is not in the index; checked Graph directly.`);
    try {
      const live = await getPage(request.pageId);
      const repaired = await repairSection(index, live.parentSection?.id, options.indexPath);
      return {
        source: 'graph',
        repaired: repaired.repaired,
        candidates: repaired.page ? [repaired.page] : [],
        page: repaired.page,
        notes,
        stats,
      };
    } catch {
      return { source: 'none', repaired: false, candidates: [], notes, stats };
    }
  }

  // 2. Title, optionally scoped by notebook/section.
  if (request.title) {
    const exact = findPageByPath(
      index,
      {
        notebookName: request.notebook ?? '',
        sectionName: request.section ?? '',
        title: request.title,
      },
    );
    if (exact.length === 1) {
      return { page: exact[0], source: 'index', repaired: false, candidates: exact, notes, stats };
    }
    if (exact.length > 1) {
      notes.push('Multiple pages share that title; returning all candidates.');
      return { source: 'index', repaired: false, candidates: exact, notes, stats };
    }

    // Fall back to a scoped substring search over the mirrors.
    const candidates = searchIndex(index, request.title, {
      notebook: request.notebook,
      section: request.section,
      limit: 25,
    });
    if (candidates.length === 1) {
      return { page: candidates[0], source: 'index', repaired: false, candidates, notes, stats };
    }
    if (candidates.length > 1) {
      notes.push(`"${request.title}" matched ${candidates.length} pages; listing candidates.`);
      return { source: 'index', repaired: false, candidates, notes, stats };
    }

    // Nothing locally — the page may be new. Refresh the most likely section.
    const sectionId = request.section
      ? index.notebooks
          .flatMap((notebook) => notebook.sections)
          .find(
            (section) =>
              section.name.toLowerCase() === request.section!.toLowerCase() &&
              (!request.notebook || matchesNotebook(index, section.id, request.notebook)),
          )?.id
      : undefined;

    if (sectionId) {
      notes.push('No local match; refreshed that section from Graph.');
      const repaired = await repairSection(index, sectionId, options.indexPath);
      return {
        source: 'graph',
        repaired: repaired.repaired,
        page: repaired.page,
        candidates: repaired.page ? [repaired.page] : [],
        notes,
        stats,
      };
    }
  }

  return { source: 'none', repaired: false, candidates: [], notes, stats };
};

const matchesNotebook = (index: PageIndex, sectionId: string, notebookName: string): boolean =>
  index.notebooks.some(
    (notebook) =>
      notebook.name.toLowerCase() === notebookName.toLowerCase() &&
      notebook.sections.some((section) => section.id === sectionId),
  );

/** Re-read one section and write the result back into the index. */
const repairSection = async (
  index: PageIndex,
  sectionId: string | undefined,
  indexPath?: string,
): Promise<{ page?: IndexedPage; repaired: boolean }> => {
  if (!sectionId) return { repaired: false };

  const owner = index.notebooks.find((notebook) =>
    notebook.sections.some((section) => section.id === sectionId),
  );
  const section = owner?.sections.find((candidate) => candidate.id === sectionId);
  if (!owner || !section) return { repaired: false };

  const pages = await listAllPagesInSection(sectionId);
  const refreshed: IndexedPage[] = pages.map((page) => ({
    id: page.id,
    title: page.title ?? '',
    notebookId: owner.id,
    notebookName: owner.name,
    sectionId,
    sectionName: section.name,
    groupPath: section.groupPath,
    createdDateTime: page.createdDateTime,
    lastModifiedDateTime: page.lastModifiedDateTime,
    webUrl: page.links?.oneNoteWebUrl?.href,
  }));

  const next: PageIndex = {
    ...index,
    refreshedAt: new Date().toISOString(),
    pages: [...index.pages.filter((page) => page.sectionId !== sectionId), ...refreshed],
  };
  await writeIndex(next, indexPath);

  return { page: refreshed.length === 1 ? refreshed[0] : undefined, repaired: true };
};

/**
 * Refresh whole sections and persist. Thin wrapper so tool handlers can keep the
 * mirror current after a write without importing the sync layer directly.
 */
export const refreshSections = async (
  sectionIds: readonly string[],
  options: { indexPath?: string; concurrency?: number } = {},
): Promise<{ refreshed: number; pages: number } | undefined> => {
  const index = await readIndex(options.indexPath);
  if (!index) return undefined;

  const next = await syncIndex(index, {
    sections: [...sectionIds],
    concurrency: options.concurrency,
  });
  await writeIndex(next, options.indexPath);
  return { refreshed: sectionIds.length, pages: next.pages.length };
};
