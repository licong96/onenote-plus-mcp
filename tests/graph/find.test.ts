import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { Page, Section } from '@/graph/types.js';

vi.mock('../../src/graph/pages.js', () => ({
  listAllPagesInSection: vi.fn(),
  getPageContent: vi.fn(),
}));

vi.mock('../../src/graph/sections.js', () => ({
  getSection: vi.fn(),
}));

vi.mock('../../src/graph/tree.js', () => ({
  collectNotebookTree: vi.fn(),
}));

import { getPageContent, listAllPagesInSection } from '@/graph/pages.js';
import { getSection } from '@/graph/sections.js';
import { collectNotebookTree } from '@/graph/tree.js';
import { findPages } from '@/graph/find.js';

const stubPage = (overrides: Partial<Page> = {}): Page => ({
  id: 'page-1',
  title: 'Test Page',
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-1/content',
  links: { oneNoteWebUrl: { href: 'https://onenote.com/page-1' } },
  ...overrides,
});

const stubSection = (overrides: Partial<Section> = {}): Section => ({
  id: 'sec-1',
  displayName: 'Test Section',
  isDefault: false,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  parentNotebook: { id: 'nb-1', displayName: 'Notebook One' },
  parentSectionGroup: null,
  ...overrides,
});

const oneNotebookTree = (
  sections: Array<{ id: string; name: string; groupPath?: string[] }>,
): unknown[] => [
  {
    id: 'nb-1',
    name: 'Notebook One',
    isDefault: true,
    isShared: false,
    lastModified: '2024-06-01T00:00:00Z',
    sections: sections.map((section) => ({
      id: section.id,
      name: section.name,
      isDefault: false,
      lastModified: '2024-06-01T00:00:00Z',
      groupPath: section.groupPath ?? [],
    })),
  },
];

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(getSection).mockResolvedValue(stubSection());
  vi.mocked(listAllPagesInSection).mockResolvedValue([]);
  vi.mocked(collectNotebookTree).mockResolvedValue([]);
});

describe('findPages scope resolution', () => {
  it('looks up a single section directly and never walks the notebook tree', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([stubPage({ title: 'Alpha' })]);

    const result = await findPages({ query: 'alpha', sectionId: 'sec-1' });

    expect(getSection).toHaveBeenCalledWith('sec-1');
    expect(collectNotebookTree).not.toHaveBeenCalled();
    expect(result.scanned).toMatchObject({ sections: 1, notebooks: 1, pages: 1 });
  });

  it('carries the parent section group into the match', async () => {
    vi.mocked(getSection).mockResolvedValue(
      stubSection({ parentSectionGroup: { id: 'sg-1', displayName: 'Group One' } }),
    );
    vi.mocked(listAllPagesInSection).mockResolvedValue([stubPage({ title: 'Alpha' })]);

    const result = await findPages({ query: 'alpha', sectionId: 'sec-1' });
    expect(result.matches[0]!.groupPath).toEqual(['Group One']);
  });

  it('resolves notebook scope through the tree and labels each match with its section', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValue(
      oneNotebookTree([
        { id: 'sec-a', name: 'Section A' },
        { id: 'sec-b', name: 'Section B', groupPath: ['Group One'] },
      ]) as never,
    );
    vi.mocked(listAllPagesInSection).mockImplementation(async (sectionId: string) =>
      sectionId === 'sec-b' ? [stubPage({ id: 'p-b', title: 'Alpha in B' })] : [],
    );

    const result = await findPages({ query: 'alpha', notebookId: 'nb-1' });

    expect(collectNotebookTree).toHaveBeenCalledWith('nb-1', expect.anything());
    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({
      section: 'Section B',
      notebook: 'Notebook One',
      groupPath: ['Group One'],
    });
  });

  it('notes the cost when no scope is supplied', async () => {
    const result = await findPages({ query: 'anything' });
    expect(result.notes.join(' ')).toContain('entire account was scanned');
  });

  it('does not add the whole-account note when a scope is given', async () => {
    const result = await findPages({ query: 'anything', notebookId: 'nb-1' });
    expect(result.notes.join(' ')).not.toContain('entire account was scanned');
  });

  it('rejects an empty query', async () => {
    await expect(findPages({ query: '   ' })).rejects.toThrow('at least one character');
  });
});

describe('findPages title matching', () => {
  it('matches titles case-insensitively and reports matchedIn=title', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'p1', title: 'Notes About ENGLISH' }),
      stubPage({ id: 'p2', title: 'Unrelated' }),
    ]);

    const result = await findPages({ query: 'english', sectionId: 'sec-1' });

    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({ id: 'p1', matchedIn: 'title' });
  });

  it('matches CJK substrings', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'p1', title: '我的英语储蓄账户目标20万' }),
    ]);

    const result = await findPages({ query: '英语', sectionId: 'sec-1' });
    expect(result.matches.map((match) => match.id)).toEqual(['p1']);
  });

  it('sorts matches newest first and applies the limit', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'old', title: 'hit old', lastModifiedDateTime: '2020-01-01T00:00:00Z' }),
      stubPage({ id: 'new', title: 'hit new', lastModifiedDateTime: '2025-01-01T00:00:00Z' }),
      stubPage({ id: 'mid', title: 'hit mid', lastModifiedDateTime: '2023-01-01T00:00:00Z' }),
    ]);

    const result = await findPages({ query: 'hit', sectionId: 'sec-1', limit: 2 });
    expect(result.matches.map((match) => match.id)).toEqual(['new', 'mid']);
  });

  it('suggests the content pass when a title search comes up empty', async () => {
    const result = await findPages({ query: 'ghost', sectionId: 'sec-1' });
    expect(result.notes.join(' ')).toContain('includeContent: true');
  });
});

describe('findPages content matching', () => {
  it('does not download page bodies unless asked', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([stubPage({ title: 'Unrelated' })]);

    const result = await findPages({ query: 'needle', sectionId: 'sec-1' });

    expect(getPageContent).not.toHaveBeenCalled();
    expect(result.scanned.contentsFetched).toBe(0);
  });

  it('finds a term inside the page body and returns a snippet', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'p1', title: 'Unrelated' }),
    ]);
    vi.mocked(getPageContent).mockResolvedValue(
      '<html><body><h1>标题</h1><p>这是一段关于英语学习的记录</p></body></html>',
    );

    const result = await findPages({ query: '英语', sectionId: 'sec-1', includeContent: true });

    expect(getPageContent).toHaveBeenCalledWith('p1');
    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]).toMatchObject({ id: 'p1', matchedIn: 'content' });
    expect(result.matches[0]!.snippet).toContain('英语');
    expect(result.matches[0]!.snippet).not.toContain('<p>');
    expect(result.scanned.contentsFetched).toBe(1);
  });

  it('strips markup and decodes entities before matching', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([stubPage({ id: 'p1', title: 'x' })]);
    vi.mocked(getPageContent).mockResolvedValue(
      '<html><head><style>p{color:red}</style></head><body><p>Tom &amp; Jerry</p></body></html>',
    );

    const result = await findPages({ query: 'tom & jerry', sectionId: 'sec-1', includeContent: true });
    expect(result.matches).toHaveLength(1);
  });

  it('never matches markup that was stripped out', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([stubPage({ id: 'p1', title: 'x' })]);
    vi.mocked(getPageContent).mockResolvedValue('<html><body><span>plain</span></body></html>');

    const result = await findPages({ query: 'span', sectionId: 'sec-1', includeContent: true });
    expect(result.matches).toEqual([]);
  });

  it('caps the content pass and reports truncation', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'p1', title: 'a' }),
      stubPage({ id: 'p2', title: 'b' }),
      stubPage({ id: 'p3', title: 'c' }),
    ]);
    vi.mocked(getPageContent).mockResolvedValue('<p>nothing here</p>');

    const result = await findPages({
      query: 'zzz',
      sectionId: 'sec-1',
      includeContent: true,
      maxContentPages: 1,
    });

    expect(getPageContent).toHaveBeenCalledTimes(1);
    expect(result.contentScanTruncated).toBe(true);
    expect(result.notes.join(' ')).toContain('stopped at 1 pages');
  });

  it('skips pages whose content cannot be read instead of failing the search', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'broken', title: 'a' }),
      stubPage({ id: 'good', title: 'b' }),
    ]);
    vi.mocked(getPageContent).mockImplementation(async (pageId: string) => {
      if (pageId === 'broken') throw new Error('410 gone');
      return '<p>the needle is here</p>';
    });

    const result = await findPages({ query: 'needle', sectionId: 'sec-1', includeContent: true });

    expect(result.matches.map((match) => match.id)).toEqual(['good']);
  });

  it('does not count a title match as a content candidate', async () => {
    vi.mocked(listAllPagesInSection).mockResolvedValue([
      stubPage({ id: 'p1', title: 'needle in title' }),
    ]);
    vi.mocked(getPageContent).mockResolvedValue('<p>irrelevant</p>');

    const result = await findPages({ query: 'needle', sectionId: 'sec-1', includeContent: true });

    expect(getPageContent).not.toHaveBeenCalled();
    expect(result.matches).toHaveLength(1);
    expect(result.matches[0]!.matchedIn).toBe('title');
  });
});
