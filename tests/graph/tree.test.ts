import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Notebook, Section, SectionGroup } from '@/graph/types.js';

vi.mock('../../src/graph/client.js', () => ({
  paginate: vi.fn(),
  graphRequest: vi.fn(),
  graphRequestRaw: vi.fn(),
  paginateUntil: vi.fn(),
}));

import { paginate } from '@/graph/client.js';
import {
  collectNotebookTree,
  collectSectionsForNotebook,
  countPagesForSections,
  countPagesInSection,
  getNotebook,
  listNotebooksRaw,
} from '@/graph/tree.js';

type PaginateLike = (path: string, options?: Record<string, unknown>) => Promise<unknown[]>;

const stubSection = (id: string, name: string, overrides: Partial<Section> = {}): Section => ({
  id,
  displayName: name,
  isDefault: false,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  ...overrides,
});

const stubGroup = (id: string, name: string): SectionGroup => ({
  id,
  displayName: name,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  parentNotebook: { id: 'nb-1', displayName: 'Notebook One' },
  parentSectionGroup: null,
});

const stubNotebook = (overrides: Partial<Notebook> = {}): Notebook => ({
  id: 'nb-1',
  displayName: 'Notebook One',
  isDefault: true,
  isShared: false,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  links: { oneNoteWebUrl: { href: 'https://onenote.com/nb-1' } },
  ...overrides,
});

/**
 * Route each paginate() call by the endpoint it targets.
 *
 * Longest fragment wins, so a specific route like
 * `/notebooks/nb-1/sections` is not swallowed by the broader
 * `/me/onenote/notebooks` collection route that also matches it.
 */
const routePaginate = (routes: Array<[string, unknown[]]>): void => {
  const sorted = [...routes].sort((left, right) => right[0].length - left[0].length);
  const implementation = (path: string): Promise<unknown[]> => {
    const match = sorted.find(([fragment]) => path.includes(fragment));
    return Promise.resolve(match ? match[1] : []);
  };
  vi.mocked(paginate).mockImplementation(implementation as unknown as typeof paginate);
};

beforeEach(() => {
  vi.clearAllMocks();
});

describe('collectSectionsForNotebook', () => {
  it('flattens top-level sections and sections nested in section groups, deepest first path', async () => {
    routePaginate([
      ['/notebooks/nb-1/sections', [stubSection('sec-top', 'Zeta Top')]],
      ['/notebooks/nb-1/sectionGroups', [stubGroup('sg-1', 'Group One')]],
      ['/sectionGroups/sg-1/sections', [stubSection('sec-in-group', 'Alpha In Group')]],
      ['/sectionGroups/sg-1/sectionGroups', [stubGroup('sg-2', 'Nested Group')]],
      ['/sectionGroups/sg-2/sections', [stubSection('sec-nested', 'Beta Nested')]],
      ['/sectionGroups/sg-2/sectionGroups', []],
    ]);

    const sections = await collectSectionsForNotebook('nb-1');

    expect(sections.map((section) => section.id)).toEqual(['sec-in-group', 'sec-nested', 'sec-top']);
    expect(sections.find((section) => section.id === 'sec-top')?.groupPath).toEqual([]);
    expect(sections.find((section) => section.id === 'sec-in-group')?.groupPath).toEqual(['Group One']);
    expect(sections.find((section) => section.id === 'sec-nested')?.groupPath).toEqual([
      'Group One',
      'Nested Group',
    ]);
  });

  it('sorts sections by name across the whole notebook', async () => {
    routePaginate([
      ['/notebooks/nb-1/sections', [stubSection('c', 'Charlie'), stubSection('a', 'Alpha')]],
      ['/notebooks/nb-1/sectionGroups', []],
    ]);

    const sections = await collectSectionsForNotebook('nb-1');
    expect(sections.map((section) => section.name)).toEqual(['Alpha', 'Charlie']);
  });

  it('maps section metadata onto the node shape', async () => {
    routePaginate([
      [
        '/notebooks/nb-1/sections',
        [
          stubSection('sec-1', 'Notes', {
            isDefault: true,
            lastModifiedDateTime: '2025-03-04T05:06:07Z',
          }),
        ],
      ],
      ['/notebooks/nb-1/sectionGroups', []],
    ]);

    const [section] = await collectSectionsForNotebook('nb-1');
    expect(section).toEqual({
      id: 'sec-1',
      name: 'Notes',
      isDefault: true,
      lastModified: '2025-03-04T05:06:07Z',
      groupPath: [],
    });
  });

  it('asks for neither $orderby nor $expand on the section collections', async () => {
    routePaginate([
      ['/notebooks/nb-1/sections', []],
      ['/notebooks/nb-1/sectionGroups', []],
    ]);

    await collectSectionsForNotebook('nb-1');

    for (const call of vi.mocked(paginate).mock.calls) {
      const query = (call[1] as { query?: Record<string, unknown> } | undefined)?.query ?? {};
      expect(query).not.toHaveProperty('$orderby');
      expect(query).not.toHaveProperty('$expand');
      expect(query.$select).toContain('displayName');
    }
  });
});

describe('countPagesInSection', () => {
  it('counts the pages returned for a section', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([{ id: 'p1' }, { id: 'p2' }, { id: 'p3' }] as never);

    await expect(countPagesInSection('sec-1')).resolves.toBe(3);

    const [path, options] = vi.mocked(paginate).mock.calls[0]!;
    expect(path).toBe('/me/onenote/sections/sec-1/pages');
    expect((options as { query: Record<string, unknown> }).query).toEqual({
      $select: 'id',
      $top: 100,
    });
  });

  it('encodes section IDs containing "!"', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([] as never);
    await countPagesInSection('0-B18ADB61AADA4FE8!267');
    expect(vi.mocked(paginate).mock.calls[0]![0]).toBe(
      '/me/onenote/sections/0-B18ADB61AADA4FE8!267/pages',
    );
  });

  it('attaches counts to every section', async () => {
    vi.mocked(paginate).mockResolvedValue([{ id: 'p' }] as never);
    const sections = [
      { id: 's1', name: 'A', isDefault: false, lastModified: 'x', groupPath: [] },
      { id: 's2', name: 'B', isDefault: false, lastModified: 'x', groupPath: [] },
    ];

    const counted = await countPagesForSections(sections);
    expect(counted.map((section) => section.pageCount)).toEqual([1, 1]);
    expect(counted.map((section) => section.name)).toEqual(['A', 'B']);
  });
});

describe('getNotebook', () => {
  it('filters the notebook collection by id', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([stubNotebook()] as never);

    const notebook = await getNotebook('nb-1');
    expect(notebook.displayName).toBe('Notebook One');

    const [path, options] = vi.mocked(paginate).mock.calls[0]!;
    expect(path).toBe('/me/onenote/notebooks');
    expect((options as { query: Record<string, unknown> }).query.$filter).toBe("id eq 'nb-1'");
  });

  it('escapes single quotes in the filter', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([stubNotebook()] as never);
    await getNotebook("nb'1");
    expect((vi.mocked(paginate).mock.calls[0]![1] as { query: Record<string, unknown> }).query.$filter).toBe(
      "id eq 'nb''1'",
    );
  });

  it('throws when the notebook does not exist', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([] as never);
    await expect(getNotebook('missing')).rejects.toThrow('Notebook not found: missing');
  });
});

describe('listNotebooksRaw', () => {
  it('does not ask for $orderby on the notebooks collection', async () => {
    vi.mocked(paginate).mockResolvedValueOnce([] as never);
    await listNotebooksRaw();
    const query = (vi.mocked(paginate).mock.calls[0]![1] as { query: Record<string, unknown> }).query;
    expect(query).not.toHaveProperty('$orderby');
  });
});

describe('collectNotebookTree', () => {
  it('scopes to a single notebook when given an id', async () => {
    routePaginate([
      ['/me/onenote/notebooks', [stubNotebook()]],
      ['/notebooks/nb-1/sections', [stubSection('sec-1', 'Notes')]],
      ['/notebooks/nb-1/sectionGroups', []],
    ]);

    const tree = await collectNotebookTree('nb-1');
    expect(tree).toHaveLength(1);
    expect(tree[0]).toMatchObject({
      id: 'nb-1',
      name: 'Notebook One',
      isDefault: true,
      isShared: false,
      webUrl: 'https://onenote.com/nb-1',
    });
    expect(tree[0]!.sections.map((section) => section.name)).toEqual(['Notes']);
  });

  it('maps every notebook when no id is given', async () => {
    routePaginate([
      [
        '/me/onenote/notebooks',
        [stubNotebook(), stubNotebook({ id: 'nb-2', displayName: 'Notebook Two', isDefault: false })],
      ],
      ['/notebooks/nb-1/sections', [stubSection('sec-1', 'Notes')]],
      ['/notebooks/nb-1/sectionGroups', []],
      ['/notebooks/nb-2/sections', []],
      ['/notebooks/nb-2/sectionGroups', []],
    ]);

    const tree = await collectNotebookTree();
    expect(tree.map((notebook) => notebook.name)).toEqual(['Notebook One', 'Notebook Two']);
    expect(tree[1]!.sections).toEqual([]);
  });

  it('omits webUrl when a notebook has no links', async () => {
    routePaginate([
      ['/me/onenote/notebooks', [stubNotebook({ links: undefined })]],
      ['/notebooks/nb-1/sections', []],
      ['/notebooks/nb-1/sectionGroups', []],
    ]);

    const tree = await collectNotebookTree();
    expect(tree[0]!.webUrl).toBeUndefined();
  });
});
