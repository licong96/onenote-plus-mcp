import { listAllPagesInSection } from '@/graph/pages.js';
import { collectNotebookTree, type NotebookNode } from '@/graph/tree.js';
import type { Page } from '@/graph/types.js';
import { DEFAULT_CONCURRENCY, mapWithConcurrency } from '@/util/concurrency.js';
import {
  INDEX_VERSION,
  indexStats,
  readIndex,
  writeIndex,
  type IndexedNotebook,
  type IndexedPage,
  type PageIndex,
} from './store.js';

/**
 * Build and maintain the on-disk mirror of the OneNote structure.
 *
 * A full rebuild is ~185 requests (13 notebooks + 172 sections + one page list
 * each) which, at the measured 2–7s per request, is measured in minutes even at
 * bounded concurrency. That cost is acceptable exactly once; afterwards the
 * per-section refresh keeps it current at a fraction of the price.
 */

const toIndexedPage = (
  page: Page,
  notebook: { id: string; name: string },
  section: { id: string; name: string; groupPath: string[] },
): IndexedPage => ({
  id: page.id,
  title: page.title ?? '',
  notebookId: notebook.id,
  notebookName: notebook.name,
  sectionId: section.id,
  sectionName: section.name,
  groupPath: section.groupPath,
  createdDateTime: page.createdDateTime,
  lastModifiedDateTime: page.lastModifiedDateTime,
  webUrl: page.links?.oneNoteWebUrl?.href,
});

const toIndexedNotebook = (notebook: NotebookNode): IndexedNotebook => ({
  id: notebook.id,
  name: notebook.name,
  isDefault: notebook.isDefault,
  isShared: notebook.isShared,
  lastModifiedDateTime: notebook.lastModified,
  sections: notebook.sections.map((section) => ({
    id: section.id,
    name: section.name,
    groupPath: section.groupPath,
    lastModifiedDateTime: section.lastModified,
  })),
});

export interface BuildOptions {
  concurrency?: number;
  /** Progress callback, for long first-time builds run in the background. */
  onProgress?: (done: number, total: number) => void;
}

/**
 * Walk the whole account and produce a complete index.
 *
 * Notebooks are collected concurrently but each notebook's pages are collected
 * inside its own worker, so concurrency stays bounded at `concurrency` requests
 * in flight overall rather than multiplying per level.
 */
export const buildIndex = async (options: BuildOptions = {}): Promise<PageIndex> => {
  const concurrency = options.concurrency ?? DEFAULT_CONCURRENCY;
  const tree = await collectNotebookTree(undefined, concurrency);

  const sections: Array<{
    notebook: { id: string; name: string };
    section: { id: string; name: string; groupPath: string[] };
  }> = [];
  for (const notebook of tree) {
    for (const section of notebook.sections) {
      sections.push({
        notebook: { id: notebook.id, name: notebook.name },
        section: { id: section.id, name: section.name, groupPath: section.groupPath },
      });
    }
  }

  let done = 0;
  const perSection = await mapWithConcurrency(sections, concurrency, async (entry) => {
    const pages = await listAllPagesInSection(entry.section.id);
    done += 1;
    options.onProgress?.(done, sections.length);
    return pages.map((page) => toIndexedPage(page, entry.notebook, entry.section));
  });

  const now = new Date().toISOString();
  return {
    version: INDEX_VERSION,
    generatedAt: now,
    refreshedAt: now,
    notebooks: tree.map(toIndexedNotebook),
    pages: perSection.flat(),
  };
};

export interface SyncOptions {
  concurrency?: number;
  /** Sections whose lastModified is newer than the index are always re-read. */
  sections?: string[];
  onProgress?: (done: number, total: number) => void;
}

/**
 * Refresh page lists for the given sections (all sections when omitted) while
 * reusing the stored notebook/section skeleton.
 *
 * Used after create/delete/move so the mirror stays truthful without paying a
 * full rebuild. Page IDs are replaced by section, which also repairs stored IDs
 * that went stale when a page was moved.
 */
export const syncIndex = async (
  index: PageIndex,
  options: SyncOptions = {},
): Promise<PageIndex> => {
  const concurrency = options.concurrency ?? DEFAULT_CONCURRENCY;
  const wanted = options.sections ? new Set(options.sections) : undefined;

  const targets = index.notebooks.flatMap((notebook) =>
    notebook.sections
      .filter((section) => !wanted || wanted.has(section.id))
      .map((section) => ({
        notebook: { id: notebook.id, name: notebook.name },
        section: { id: section.id, name: section.name, groupPath: section.groupPath },
      })),
  );

  let done = 0;
  const refreshed = await mapWithConcurrency(targets, concurrency, async (entry) => {
    const pages = await listAllPagesInSection(entry.section.id);
    done += 1;
    options.onProgress?.(done, targets.length);
    return pages.map((page) => toIndexedPage(page, entry.notebook, entry.section));
  });

  // Wholesale replacement of the touched sections: a page that disappeared from
  // its listing must disappear from the index too.
  const touched = new Set(targets.map((entry) => entry.section.id));
  const kept = index.pages.filter((page) => !touched.has(page.sectionId));

  return {
    ...index,
    refreshedAt: new Date().toISOString(),
    pages: [...kept, ...refreshed.flat()],
  };
};

/**
 * Rebuild the skeleton (notebooks/sections) only, keeping existing pages for
 * sections that still exist. Cheaper than a full build when sections were
 * added, renamed, or removed but page data is presumed unchanged.
 */
export const rebuildSkeleton = async (
  index: PageIndex,
  options: BuildOptions = {},
): Promise<PageIndex> => {
  const tree = await collectNotebookTree(undefined, options.concurrency ?? DEFAULT_CONCURRENCY);
  const liveSections = new Set(tree.flatMap((n) => n.sections.map((s) => s.id)));
  const nameBySection = new Map<string, { notebook: string; section: string; groupPath: string[] }>();
  for (const notebook of tree) {
    for (const section of notebook.sections) {
      nameBySection.set(section.id, {
        notebook: notebook.name,
        section: section.name,
        groupPath: section.groupPath,
      });
    }
  }

  const pages = index.pages
    // Sections that no longer exist take their pages with them.
    .filter((page) => liveSections.has(page.sectionId))
    .map((page) => {
      const names = nameBySection.get(page.sectionId);
      // Section/notebook renames must be reflected on the pages that live there.
      return names
        ? {
            ...page,
            notebookName: names.notebook,
            sectionName: names.section,
            groupPath: names.groupPath,
          }
        : page;
    });

  return {
    version: INDEX_VERSION,
    generatedAt: new Date().toISOString(),
    refreshedAt: index.refreshedAt,
    notebooks: tree.map(toIndexedNotebook),
    pages,
  };
};

export interface EnsureResult {
  index: PageIndex;
  rebuilt: boolean;
  stats: ReturnType<typeof indexStats>;
}

/**
 * Load the index, rebuilding from scratch when absent or unusable.
 * Callers that want the cache-only path should use `readIndex` directly.
 */
export const ensureIndex = async (
  options: BuildOptions & { indexPath?: string } = {},
): Promise<EnsureResult> => {
  const existing = await readIndex(options.indexPath);
  if (existing) {
    return { index: existing, rebuilt: false, stats: indexStats(existing) };
  }
  const built = await buildIndex(options);
  await writeIndex(built, options.indexPath);
  return { index: built, rebuilt: true, stats: indexStats(built) };
};
