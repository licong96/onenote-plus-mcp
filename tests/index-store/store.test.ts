import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  INDEX_VERSION,
  buildPathMap,
  emptyIndex,
  findPageByPath,
  indexStats,
  readIndex,
  removePages,
  searchIndex,
  upsertPages,
  writeIndex,
  type IndexedPage,
  type PageIndex,
} from '@/index-store/store.js';

const page = (overrides: Partial<IndexedPage> = {}): IndexedPage => ({
  id: 'p1',
  title: 'Notes',
  notebookId: 'nb-1',
  notebookName: 'Notebook One',
  sectionId: 'sec-1',
  sectionName: 'Section A',
  groupPath: [],
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  ...overrides,
});

const indexWith = (pages: IndexedPage[]): PageIndex => ({
  version: INDEX_VERSION,
  generatedAt: '2024-06-01T00:00:00Z',
  notebooks: [
    {
      id: 'nb-1',
      name: 'Notebook One',
      isDefault: true,
      isShared: false,
      lastModifiedDateTime: '2024-06-01T00:00:00Z',
      sections: [
        { id: 'sec-1', name: 'Section A', groupPath: [], lastModifiedDateTime: '2024-06-01T00:00:00Z' },
      ],
    },
  ],
  pages,
});

let tempRoot: string;
let indexPath: string;

beforeEach(async () => {
  tempRoot = await mkdtemp(join(tmpdir(), 'onenote-index-'));
  indexPath = join(tempRoot, 'index.json');
});

afterEach(async () => {
  await rm(tempRoot, { recursive: true, force: true });
});

describe('readIndex / writeIndex', () => {
  it('round-trips an index through disk', async () => {
    const index = indexWith([page()]);
    await writeIndex(index, indexPath);

    const loaded = await readIndex(indexPath);
    expect(loaded?.pages).toEqual([page()]);
    expect(loaded?.version).toBe(INDEX_VERSION);
  });

  it('returns undefined when the file does not exist', async () => {
    await expect(readIndex(indexPath)).resolves.toBeUndefined();
  });

  it('returns undefined for a corrupt file instead of throwing', async () => {
    await writeIndex(indexWith([]), indexPath);
    const { writeFile } = await import('node:fs/promises');
    await writeFile(indexPath, '{ not json');

    await expect(readIndex(indexPath)).resolves.toBeUndefined();
  });

  it('treats a future version as unusable rather than partially usable', async () => {
    await writeIndex({ ...indexWith([page()]), version: INDEX_VERSION + 1 }, indexPath);
    await expect(readIndex(indexPath)).resolves.toBeUndefined();
  });

  it('treats an empty file as absent', async () => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(indexPath, '   ');
    await expect(readIndex(indexPath)).resolves.toBeUndefined();
  });

  it('leaves no temp file behind', async () => {
    await writeIndex(indexWith([page()]), indexPath);
    const { readdir } = await import('node:fs/promises');
    const entries = await readdir(tempRoot);
    expect(entries).toEqual(['index.json']);
  });
});

describe('indexStats', () => {
  it('counts notebooks, sections, and pages', () => {
    expect(indexStats(indexWith([page(), page({ id: 'p2' })]))).toEqual({
      notebooks: 1,
      sections: 1,
      pages: 2,
    });
  });

  it('reports zeroes for an empty index', () => {
    expect(indexStats(emptyIndex())).toEqual({ notebooks: 0, sections: 0, pages: 0 });
  });
});

describe('findPageByPath', () => {
  it('finds a page by notebook + section + title, case-insensitively', () => {
    const index = indexWith([page()]);
    const found = findPageByPath(index, {
      notebookName: 'notebook one',
      sectionName: 'section a',
      title: 'NOTES',
    });
    expect(found.map((entry) => entry.id)).toEqual(['p1']);
  });

  it('returns every page sharing a title', () => {
    const index = indexWith([page(), page({ id: 'p2' })]);
    expect(findPageByPath(index, {
      notebookName: 'Notebook One',
      sectionName: 'Section A',
      title: 'Notes',
    })).toHaveLength(2);
  });

  it('does not match across notebooks with the same section and title', () => {
    const index = indexWith([page()]);
    expect(
      findPageByPath(index, {
        notebookName: 'Other Notebook',
        sectionName: 'Section A',
        title: 'Notes',
      }),
    ).toEqual([]);
  });

  it('buildPathMap groups by the composite key', () => {
    const map = buildPathMap([page(), page({ id: 'p2', title: 'Other' })]);
    expect(map.size).toBe(2);
  });
});

describe('searchIndex', () => {
  it('matches titles case-insensitively', () => {
    const index = indexWith([page({ title: 'About ENGLISH' })]);
    expect(searchIndex(index, 'english').map((entry) => entry.id)).toEqual(['p1']);
  });

  it('matches CJK substrings', () => {
    const index = indexWith([page({ title: '我的英语储蓄账户' })]);
    expect(searchIndex(index, '英语')).toHaveLength(1);
  });

  it('matches section and notebook names, not just titles', () => {
    const index = indexWith([page({ title: 'Unrelated' })]);
    expect(searchIndex(index, 'Section A')).toHaveLength(1);
    expect(searchIndex(index, 'Notebook One')).toHaveLength(1);
  });

  it('matches section-group paths', () => {
    const index = indexWith([page({ title: 'Unrelated', groupPath: ['Archive'] })]);
    expect(searchIndex(index, 'archive')).toHaveLength(1);
  });

  it('ranks title hits above structural hits', () => {
    const index = indexWith([
      page({ id: 'structural', title: 'Unrelated', sectionName: 'Archive' }),
      page({ id: 'title-hit', title: 'Archive' }),
    ]);
    expect(searchIndex(index, 'archive')[0]!.id).toBe('title-hit');
  });

  it('ranks newer pages first within the same tier', () => {
    const index = indexWith([
      page({ id: 'old', title: 'hit', lastModifiedDateTime: '2020-01-01T00:00:00Z' }),
      page({ id: 'new', title: 'hit', lastModifiedDateTime: '2025-01-01T00:00:00Z' }),
    ]);
    expect(searchIndex(index, 'hit').map((entry) => entry.id)).toEqual(['new', 'old']);
  });

  it('applies notebook and section filters', () => {
    const index = indexWith([
      page({ id: 'p1', title: 'hit', sectionName: 'Section A' }),
      page({ id: 'p2', title: 'hit', sectionName: 'Section B' }),
    ]);
    expect(searchIndex(index, 'target', { section: 'Section B' }).map((e) => e.id)).toEqual([]);
    expect(searchIndex(index, 'hit', { section: 'Section B' }).map((e) => e.id)).toEqual(['p2']);
    expect(searchIndex(index, 'hit', { section: 'Nope' })).toEqual([]);
  });

  it('honours the limit', () => {
    const index = indexWith([page({ id: 'a', title: 'x' }), page({ id: 'b', title: 'x' })]);
    expect(searchIndex(index, 'x', { limit: 1 })).toHaveLength(1);
  });

  it('returns nothing for an empty query', () => {
    expect(searchIndex(indexWith([page()]), '   ')).toEqual([]);
  });
});

describe('removePages / upsertPages', () => {
  it('removes the named pages and leaves the rest', () => {
    const next = removePages(indexWith([page(), page({ id: 'p2' })]), ['p1']);
    expect(next.pages.map((entry) => entry.id)).toEqual(['p2']);
  });

  it('does not mutate the input index', () => {
    const original = indexWith([page()]);
    removePages(original, ['p1']);
    expect(original.pages).toHaveLength(1);
  });

  it('upserts by id, replacing an existing entry', () => {
    const next = upsertPages(indexWith([page()]), [page({ title: 'Renamed' })]);
    expect(next.pages).toHaveLength(1);
    expect(next.pages[0]!.title).toBe('Renamed');
  });

  it('appends genuinely new pages', () => {
    const next = upsertPages(indexWith([page()]), [page({ id: 'p2' })]);
    expect(next.pages.map((entry) => entry.id).sort()).toEqual(['p1', 'p2']);
  });
});
