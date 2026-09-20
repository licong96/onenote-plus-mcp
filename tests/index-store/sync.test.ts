import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Page } from '@/graph/types.js';
import { INDEX_VERSION, type PageIndex } from '@/index-store/store.js';

vi.mock('../../src/graph/pages.js', () => ({
  listAllPagesInSection: vi.fn(),
  getPageContent: vi.fn(),
  getPage: vi.fn(),
}));

vi.mock('../../src/graph/sections.js', () => ({
  getSection: vi.fn(),
}));

vi.mock('../../src/graph/tree.js', () => ({
  collectNotebookTree: vi.fn(),
}));

import { listAllPagesInSection } from '@/graph/pages.js';
import { getSection } from '@/graph/sections.js';
import { collectNotebookTree } from '@/graph/tree.js';
import { buildIndex, rebuildSkeleton, syncIndex } from '@/index-store/sync.js';

const stubPage = (id: string, title: string, overrides: Partial<Page> = {}): Page => ({
  id,
  title,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
  links: { oneNoteWebUrl: { href: `https://onenote.com/${id}` } },
  ...overrides,
});

const sectionNode = (id: string, name: string, groupPath: string[] = []) => ({
  id,
  name,
  isDefault: false,
  lastModified: '2024-06-01T00:00:00Z',
  groupPath,
});

const notebookNode = (
  id: string,
  name: string,
  sections: ReturnType<typeof sectionNode>[],
) => ({
  id,
  name,
  isDefault: false,
  isShared: false,
  lastModified: '2024-06-01T00:00:00Z',
  sections,
});

const indexOf = (pages: PageIndex['pages'], notebooks: PageIndex['notebooks'] = []): PageIndex => ({
  version: INDEX_VERSION,
  generatedAt: '2024-06-01T00:00:00Z',
  notebooks,
  pages,
});

beforeEach(() => {
  vi.clearAllMocks();
});

describe('buildIndex', () => {
  it('flattens the tree into notebook/section/page records', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'Notebook One', [
        sectionNode('sec-1', 'Section A'),
        sectionNode('sec-2', 'Section B', ['Group One']),
      ]),
    ] as never);
    vi.mocked(listAllPagesInSection).mockImplementation(async (sectionId: string) =>
      sectionId === 'sec-1'
        ? [stubPage('p1', 'First')]
        : [stubPage('p2', 'Second')],
    );

    const index = await buildIndex();

    expect(index.version).toBe(INDEX_VERSION);
    expect(index.pages).toHaveLength(2);
    expect(index.pages.find((page) => page.id === 'p2')).toMatchObject({
      notebookId: 'nb-1',
      notebookName: 'Notebook One',
      sectionId: 'sec-2',
      sectionName: 'Section B',
      groupPath: ['Group One'],
    });
    expect(index.notebooks[0]!.sections).toHaveLength(2);
  });

  it('records the page web url and timestamps', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'NB', [sectionNode('sec-1', 'S')]),
    ] as never);
    vi.mocked(listAllPagesInSection).mockResolvedValueOnce([
      stubPage('p1', 'T', { createdDateTime: '2023-01-01T00:00:00Z' }),
    ]);

    const index = await buildIndex();
    expect(index.pages[0]).toMatchObject({
      createdDateTime: '2023-01-01T00:00:00Z',
      webUrl: 'https://onenote.com/p1',
    });
  });

  it('reports progress once per section', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'NB', [sectionNode('s1', 'A'), sectionNode('s2', 'B')]),
    ] as never);
    vi.mocked(listAllPagesInSection).mockResolvedValue([]);
    const onProgress = vi.fn();

    await buildIndex({ onProgress });
    expect(onProgress).toHaveBeenCalledTimes(2);
    expect(onProgress).toHaveBeenLastCalledWith(2, 2);
  });

  it('produces an index with no pages when every section is empty', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'NB', [sectionNode('s1', 'A')]),
    ] as never);
    vi.mocked(listAllPagesInSection).mockResolvedValue([]);

    const index = await buildIndex();
    expect(index.pages).toEqual([]);
    expect(index.notebooks).toHaveLength(1);
  });
});

describe('syncIndex', () => {
  it('refreshes only the named sections', async () => {
    const index = indexOf(
      [
        {
          id: 'p1',
          title: 'Old',
          notebookId: 'nb-1',
          notebookName: 'NB',
          sectionId: 'sec-1',
          sectionName: 'A',
          groupPath: [],
          createdDateTime: 'x',
          lastModifiedDateTime: 'x',
        },
      ],
      [
        {
          id: 'nb-1',
          name: 'NB',
          isDefault: false,
          isShared: false,
          lastModifiedDateTime: 'x',
          sections: [
            { id: 'sec-1', name: 'A', groupPath: [], lastModifiedDateTime: 'x' },
            { id: 'sec-2', name: 'B', groupPath: [], lastModifiedDateTime: 'x' },
          ],
        },
      ],
    );
    vi.mocked(listAllPagesInSection).mockResolvedValueOnce([stubPage('p1', 'Renamed')]);

    const next = await syncIndex(index, { sections: ['sec-1'] });

    expect(listAllPagesInSection).toHaveBeenCalledTimes(1);
    expect(listAllPagesInSection).toHaveBeenCalledWith('sec-1');
    expect(next.pages[0]!.title).toBe('Renamed');
  });

  it('drops pages that vanished from a refreshed section', async () => {
    const index = indexOf(
      [
        {
          id: 'gone',
          title: 'Deleted',
          notebookId: 'nb-1',
          notebookName: 'NB',
          sectionId: 'sec-1',
          sectionName: 'A',
          groupPath: [],
          createdDateTime: 'x',
          lastModifiedDateTime: 'x',
        },
      ],
      [
        {
          id: 'nb-1',
          name: 'NB',
          isDefault: false,
          isShared: false,
          lastModifiedDateTime: 'x',
          sections: [{ id: 'sec-1', name: 'A', groupPath: [], lastModifiedDateTime: 'x' }],
        },
      ],
    );
    vi.mocked(listAllPagesInSection).mockResolvedValueOnce([]);

    const next = await syncIndex(index, { sections: ['sec-1'] });
    expect(next.pages).toEqual([]);
  });

  it('refreshes every section when none are named', async () => {
    const index = indexOf(
      [],
      [
        {
          id: 'nb-1',
          name: 'NB',
          isDefault: false,
          isShared: false,
          lastModifiedDateTime: 'x',
          sections: [
            { id: 's1', name: 'A', groupPath: [], lastModifiedDateTime: 'x' },
            { id: 's2', name: 'B', groupPath: [], lastModifiedDateTime: 'x' },
          ],
        },
      ],
    );
    vi.mocked(listAllPagesInSection).mockResolvedValue([]);

    await syncIndex(index);
    expect(listAllPagesInSection).toHaveBeenCalledTimes(2);
  });

  it('stamps refreshedAt', async () => {
    const index = indexOf([], [
      {
        id: 'nb-1',
        name: 'NB',
        isDefault: false,
        isShared: false,
        lastModifiedDateTime: 'x',
        sections: [{ id: 's1', name: 'A', groupPath: [], lastModifiedDateTime: 'x' }],
      },
    ]);
    vi.mocked(listAllPagesInSection).mockResolvedValue([]);

    const next = await syncIndex(index);
    expect(next.refreshedAt).toBeDefined();
    expect(next.generatedAt).toBe(index.generatedAt);
  });
});

describe('rebuildSkeleton', () => {
  it('drops pages belonging to sections that no longer exist', async () => {
    const index = indexOf(
      [
        {
          id: 'keep',
          title: 'Keep',
          notebookId: 'nb-1',
          notebookName: 'NB',
          sectionId: 'sec-live',
          sectionName: 'A',
          groupPath: [],
          createdDateTime: 'x',
          lastModifiedDateTime: 'x',
        },
        {
          id: 'orphan',
          title: 'Orphan',
          notebookId: 'nb-1',
          notebookName: 'NB',
          sectionId: 'sec-dead',
          sectionName: 'Dead',
          groupPath: [],
          createdDateTime: 'x',
          lastModifiedDateTime: 'x',
        },
      ],
      [],
    );
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'NB', [sectionNode('sec-live', 'A')]),
    ] as never);

    const next = await rebuildSkeleton(index);
    expect(next.pages.map((page) => page.id)).toEqual(['keep']);
  });

  it('propagates a section rename onto its pages', async () => {
    const index = indexOf(
      [
        {
          id: 'p1',
          title: 'T',
          notebookId: 'nb-1',
          notebookName: 'Old Notebook',
          sectionId: 'sec-1',
          sectionName: 'Old Section',
          groupPath: [],
          createdDateTime: 'x',
          lastModifiedDateTime: 'x',
        },
      ],
      [],
    );
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'New Notebook', [sectionNode('sec-1', 'New Section', ['Group'])]) as never,
    ] as never);

    const next = await rebuildSkeleton(index);
    expect(next.pages[0]).toMatchObject({
      notebookName: 'New Notebook',
      sectionName: 'New Section',
      groupPath: ['Group'],
    });
  });

  it('does not fetch page lists', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([] as never);
    await rebuildSkeleton(indexOf([]));
    expect(listAllPagesInSection).not.toHaveBeenCalled();
  });
});

describe('getSection is not called during a build', () => {
  it('relies on the tree for section metadata', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      notebookNode('nb-1', 'NB', [sectionNode('s1', 'A')]),
    ] as never);
    vi.mocked(listAllPagesInSection).mockResolvedValue([]);

    await buildIndex();
    expect(getSection).not.toHaveBeenCalled();
  });
});
