import { beforeEach, describe, expect, it, vi } from 'vitest';
import { INDEX_VERSION, type PageIndex } from '@/index-store/store.js';

vi.mock('../../src/graph/pages.js', () => ({
  listAllPagesInSection: vi.fn(),
  getPage: vi.fn(),
  getPageContent: vi.fn(),
}));

vi.mock('../../src/graph/sections.js', () => ({
  getSection: vi.fn(),
}));

vi.mock('../../src/index-store/store.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@/index-store/store.js')>();
  return { ...actual, readIndex: vi.fn(), writeIndex: vi.fn() };
});

import { getPage, listAllPagesInSection } from '@/graph/pages.js';
import { readIndex, writeIndex, type IndexedPage, type PageIndex } from '@/index-store/store.js';
import type { Page } from '@/graph/types.js';
import { resolvePage } from '@/index-store/resolve.js';

const indexed = (overrides: Partial<IndexedPage> = {}): IndexedPage => ({
  id: 'p1',
  title: '香港',
  notebookId: 'nb-1',
  notebookName: '去做的事',
  sectionId: 'sec-1',
  sectionName: '今日任务',
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
      name: '去做的事',
      isDefault: false,
      isShared: false,
      lastModifiedDateTime: 'x',
      sections: [
        { id: 'sec-1', name: '今日任务', groupPath: [], lastModifiedDateTime: 'x' },
      ],
    },
  ],
  pages,
});

const stubPage = (id: string, title: string): Page => ({
  id,
  title,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
  links: { oneNoteWebUrl: { href: `https://onenote.com/${id}` } },
});

beforeEach(() => {
  vi.clearAllMocks();
});

describe('resolvePage without an index', () => {
  it('says so instead of guessing', async () => {
    vi.mocked(readIndex).mockResolvedValue(undefined);

    const result = await resolvePage({ title: '香港' });

    expect(result.source).toBe('none');
    expect(result.notes.join(' ')).toContain('No index on disk');
  });
});

describe('resolvePage by id', () => {
  it('serves an exact id from the index without touching Graph', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([indexed()]));

    const result = await resolvePage({ pageId: 'p1' });

    expect(result.page?.id).toBe('p1');
    expect(result.source).toBe('index');
    expect(getPage).not.toHaveBeenCalled();
  });

  it('verifies an unknown id against Graph and records the answer', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));
    // A page stored by a stale index: the id is real but the mirror lacks it,
    // so Graph is consulted and the section is re-read into the index.
    vi.mocked(getPage).mockResolvedValueOnce({
      ...stubPage('sec-1-page', 'From Graph'),
      parentSection: { id: 'sec-1', displayName: '今日任务' },
    });
    vi.mocked(listAllPagesInSection).mockResolvedValueOnce([stubPage('sec-1-page', 'From Graph')]);
    vi.mocked(writeIndex).mockResolvedValue(undefined);

    const result = await resolvePage({ pageId: 'sec-1-page' });

    expect(getPage).toHaveBeenCalledWith('sec-1-page');
    expect(result.source).toBe('graph');
    expect(result.notes.join(' ')).toContain('not in the index');
    expect(writeIndex).toHaveBeenCalled();
  });

  it('reports a miss when the indexed page belongs to no known section', async () => {
    // The stored id is absent and the page's section is unknown to the index,
    // so there is nothing to refresh — the resolver must not invent an answer.
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));
    vi.mocked(getPage).mockResolvedValueOnce({
      ...stubPage('p9', 'Orphan'),
      parentSection: { id: 'sec-unknown', displayName: 'Unknown' },
    });

    const result = await resolvePage({ pageId: 'p9' });

    expect(result.page).toBeUndefined();
    expect(result.repaired).toBe(false);
    expect(writeIndex).not.toHaveBeenCalled();
  });
});

describe('resolvePage by title', () => {
  it('resolves a unique title from the index', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([indexed()]));

    const result = await resolvePage({
      title: '香港',
      notebook: '去做的事',
      section: '今日任务',
    });

    expect(result.page?.id).toBe('p1');
    expect(result.source).toBe('index');
  });

  it('returns every candidate when a title is ambiguous', async () => {
    vi.mocked(readIndex).mockResolvedValue(
      indexWith([
        indexed({ id: 'a' }),
        indexed({ id: 'b', sectionId: 'sec-1' }),
      ]),
    );

    const result = await resolvePage({
      title: '香港',
      notebook: '去做的事',
      section: '今日任务',
    });

    expect(result.candidates).toHaveLength(2);
    expect(result.notes.join(' ')).toContain('Multiple pages');
  });

  it('falls back to a scoped substring search', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([indexed({ title: '关于香港的笔记' })]));

    const result = await resolvePage({ title: '香港', notebook: '去做的事' });

    expect(result.page?.title).toBe('关于香港的笔记');
  });

  it('refreshes the named section when nothing matches locally', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));
    vi.mocked(listAllPagesInSection).mockResolvedValueOnce([stubPage('new', '新页面')]);
    vi.mocked(writeIndex).mockResolvedValue(undefined);

    const result = await resolvePage({
      title: '新页面',
      notebook: '去做的事',
      section: '今日任务',
    });

    expect(listAllPagesInSection).toHaveBeenCalledWith('sec-1');
    expect(result.source).toBe('graph');
    expect(result.repaired).toBe(true);
    expect(result.page?.title).toBe('新页面');
  });

  it('reports none when there is no section to refresh', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));

    const result = await resolvePage({ title: '不存在' });

    expect(result.source).toBe('none');
    expect(listAllPagesInSection).not.toHaveBeenCalled();
  });
});
