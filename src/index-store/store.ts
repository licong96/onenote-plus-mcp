import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { getConfigDir } from '@/config.js';

/**
 * Local index of the OneNote structure.
 *
 * Every read against Graph costs 2–7s of round-trip regardless of payload
 * size (measured: listing 7 pages took 7s), while the notebook/section skeleton
 * is essentially static and even page lists only change when the user edits.
 * So the whole structure is mirrored to one JSON file on disk and served from
 * there; Graph is only consulted when the mirror is stale or misses.
 *
 * Two index keys, deliberately:
 *   - byId: exact, O(1), used once a page id is known.
 *   - byPath: notebook + section + title, used because page IDs are NOT stable.
 *     Copying or moving a page between sections mints a new ID, so a stored ID
 *     can silently 404. The path key is the fallback that reconstructs it.
 */
export const INDEX_VERSION = 1;

export interface IndexedPage {
  id: string;
  title: string;
  notebookId: string;
  notebookName: string;
  sectionId: string;
  sectionName: string;
  /** Section-group names above the section, outermost first. */
  groupPath: string[];
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl?: string;
}

export interface IndexedSection {
  id: string;
  name: string;
  groupPath: string[];
  lastModifiedDateTime: string;
}

export interface IndexedNotebook {
  id: string;
  name: string;
  isDefault: boolean;
  isShared: boolean;
  lastModifiedDateTime: string;
  sections: IndexedSection[];
}

export interface PageIndex {
  version: number;
  /** ISO timestamp of the last full rebuild or successful sync. */
  generatedAt: string;
  /** ISO timestamp of the last per-section refresh. */
  refreshedAt?: string;
  notebooks: IndexedNotebook[];
  pages: IndexedPage[];
}

export interface PageIndexStats {
  notebooks: number;
  sections: number;
  pages: number;
}

export const emptyIndex = (): PageIndex => ({
  version: INDEX_VERSION,
  generatedAt: new Date(0).toISOString(),
  notebooks: [],
  pages: [],
});

export const indexStats = (index: PageIndex): PageIndexStats => ({
  notebooks: index.notebooks.length,
  sections: index.notebooks.reduce((sum, notebook) => sum + notebook.sections.length, 0),
  pages: index.pages.length,
});

export const getIndexPath = (): string => join(getConfigDir(), 'index.json');

export const readIndex = async (path: string = getIndexPath()): Promise<PageIndex | undefined> => {
  try {
    const raw = await readFile(path, 'utf8');
    if (raw.trim().length === 0) return undefined;
    const parsed = JSON.parse(raw) as PageIndex;
    // A future format is unusable, not partially usable — force a rebuild.
    if (parsed.version !== INDEX_VERSION) return undefined;
    return parsed;
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
    // A corrupt index must never break the server; treat it as absent.
    return undefined;
  }
};

export const indexAgeMs = async (path: string = getIndexPath()): Promise<number | undefined> => {
  try {
    const info = await stat(path);
    return Date.now() - info.mtimeMs;
  } catch {
    return undefined;
  }
};

/**
 * Write via a temp file + rename so a crash mid-write can never leave a
 * half-serialized index behind (rename is atomic within a filesystem).
 */
export const writeIndex = async (
  index: PageIndex,
  path: string = getIndexPath(),
): Promise<void> => {
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const temp = `${path}.${process.pid}.tmp`;
  await writeFile(temp, `${JSON.stringify(index, null, 2)}\n`, { encoding: 'utf8', mode: 0o600 });
  await rename(temp, path);
};

const pathKey = (notebookName: string, sectionName: string, title: string): string =>
  `${notebookName}\u0000${sectionName}\u0000${title}`.toLowerCase();

export interface PageLookup {
  notebookName: string;
  sectionName: string;
  title: string;
}

/**
 * Rebuild the path index from a page list. Built fresh on every lookup rather
 * than persisted, so a stale file can never answer a path query with data that
 * disagrees with the pages it holds.
 */
export const buildPathMap = (pages: readonly IndexedPage[]): Map<string, IndexedPage[]> => {
  const map = new Map<string, IndexedPage[]>();
  for (const page of pages) {
    const key = pathKey(page.notebookName, page.sectionName, page.title);
    const bucket = map.get(key);
    if (bucket) bucket.push(page);
    else map.set(key, [page]);
  }
  return map;
};

export const findPageByPath = (index: PageIndex, lookup: PageLookup): IndexedPage[] => {
  const key = pathKey(lookup.notebookName, lookup.sectionName, lookup.title);
  return buildPathMap(index.pages).get(key) ?? [];
};

const matches = (haystack: string, needle: string): boolean =>
  haystack.toLowerCase().includes(needle);

/**
 * Compare a haystack against a caller-supplied filter.
 *
 * The filter arrives as raw user text ("Section B", "记录我的生活") while the
 * query needle is already lowercased, so the filter must be lowercased too —
 * otherwise a capitalised filter silently matches nothing while a lowercase one
 * works. That inconsistency is worse than being uniformly case-sensitive.
 */
const matchesFilter = (haystack: string, filter: string): boolean =>
  matches(haystack, filter.trim().toLowerCase());

export interface SearchOptions {
  notebook?: string;
  section?: string;
  limit?: number;
}

/**
 * Offline fuzzy-ish search over the mirrored index: case-insensitive substring
 * match across title, section, notebook, and group path. Costs nothing on the
 * network, which is the whole point — the Graph full-text search endpoint is
 * rejected for this account.
 */
export const searchIndex = (
  index: PageIndex,
  query: string,
  options: SearchOptions = {},
): IndexedPage[] => {
  const needle = query.trim().toLowerCase();
  if (needle.length === 0) return [];

  const results = index.pages.filter((page) => {
    if (options.notebook && !matchesFilter(page.notebookName, options.notebook)) return false;
    if (options.section && !matchesFilter(page.sectionName, options.section)) return false;
    return (
      matches(page.title, needle) ||
      matches(page.sectionName, needle) ||
      matches(page.notebookName, needle) ||
      page.groupPath.some((group) => matches(group, needle))
    );
  });

  // Title hits outrank structural hits; newer pages outrank older ones.
  results.sort((left, right) => {
    const leftTitle = matches(left.title, needle) ? 0 : 1;
    const rightTitle = matches(right.title, needle) ? 0 : 1;
    if (leftTitle !== rightTitle) return leftTitle - rightTitle;
    return right.lastModifiedDateTime.localeCompare(left.lastModifiedDateTime);
  });

  return results.slice(0, Math.max(1, options.limit ?? 25));
};

export const removePages = (index: PageIndex, pageIds: readonly string[]): PageIndex => {
  const doomed = new Set(pageIds);
  return { ...index, pages: index.pages.filter((page) => !doomed.has(page.id)) };
};

export const upsertPages = (index: PageIndex, incoming: readonly IndexedPage[]): PageIndex => {
  const byId = new Map(index.pages.map((page) => [page.id, page]));
  for (const page of incoming) byId.set(page.id, page);
  return { ...index, pages: [...byId.values()] };
};
