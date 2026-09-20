import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Page, Section } from '@/graph/types.js';

vi.mock('../../src/graph/pages.js', () => ({
  listPagesInSection: vi.fn(),
  listAllPagesInSection: vi.fn(),
  copyPageToSection: vi.fn(),
  getPage: vi.fn(),
}));

vi.mock('../../src/graph/sections.js', () => ({
  getSection: vi.fn(),
}));

vi.mock('../../src/graph/tree.js', () => ({
  collectSectionsForNotebook: vi.fn(),
  collectNotebookTree: vi.fn(),
  countPagesForSections: vi.fn(),
}));

vi.mock('../../src/graph/find.js', () => ({
  findPages: vi.fn(),
}));

vi.mock('../../src/auth/index.js', () => ({
  getAuthSnapshot: vi.fn(),
}));

import { copyPageToSection, getPage, listPagesInSection } from '@/graph/pages.js';
import { getSection } from '@/graph/sections.js';
import {
  collectNotebookTree,
  collectSectionsForNotebook,
  countPagesForSections,
} from '@/graph/tree.js';
import { findPages } from '@/graph/find.js';
import { getAuthSnapshot } from '@/auth/index.js';

type ToolResult = { content: { type: string; text: string }[] };
type ToolHandler = (input: Record<string, unknown>) => Promise<ToolResult>;

const captureHandler = (registerFn: (server: McpServer) => void): ToolHandler => {
  let handler: ToolHandler | undefined;
  const mockServer = {
    registerTool: (_name: string, _meta: unknown, cb: ToolHandler) => {
      handler = cb;
    },
  } as unknown as McpServer;
  registerFn(mockServer);
  return handler!;
};

const parseResult = <T = Record<string, unknown>>(result: ToolResult): T =>
  JSON.parse(result.content[0]!.text) as T;

const stubPage = (id: string, title: string, overrides: Partial<Page> = {}): Page => ({
  id,
  title,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
  links: { oneNoteWebUrl: { href: `https://onenote.com/${id}` } },
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

const sectionNode = (id: string, name: string, groupPath: string[] = []) => ({
  id,
  name,
  isDefault: false,
  lastModified: '2024-06-01T00:00:00Z',
  groupPath,
});

const originalEnv = {
  XDG_CONFIG_HOME: process.env.XDG_CONFIG_HOME,
  ONENOTE_MCP_CLIENT_ID: process.env.ONENOTE_MCP_CLIENT_ID,
  ONENOTE_MCP_CONFIG_DIR: process.env.ONENOTE_MCP_CONFIG_DIR,
};
let tempRoot: string;

beforeEach(async () => {
  vi.clearAllMocks();
  tempRoot = await mkdtemp(join(tmpdir(), 'onenote-plus-tools-'));
  process.env.XDG_CONFIG_HOME = tempRoot;
  process.env.ONENOTE_MCP_CLIENT_ID = 'test-client-id';
  delete process.env.ONENOTE_MCP_CONFIG_DIR;
});

afterEach(async () => {
  await rm(tempRoot, { recursive: true, force: true });
  for (const [key, value] of Object.entries(originalEnv)) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
});

describe('list_pages handler', () => {
  it('lists a section and echoes the scope', async () => {
    vi.mocked(listPagesInSection).mockResolvedValueOnce([
      stubPage('p1', 'First'),
      stubPage('p2', 'Second'),
    ]);

    const { register } = await import('@/tools/listPages.js');
    const result = parseResult(await captureHandler(register)({ sectionId: 'sec-1', limit: 5 }));

    expect(listPagesInSection).toHaveBeenCalledWith('sec-1', {
      limit: 5,
      orderBy: 'lastModified',
      order: 'desc',
    });
    expect(result.scope).toEqual({ sectionId: 'sec-1' });
    expect(result.returned).toBe(2);
    expect((result.pages as Record<string, unknown>[])[0]).toMatchObject({
      id: 'p1',
      title: 'First',
      webUrl: 'https://onenote.com/p1',
    });
  });

  it('defaults to 50 rows, newest first', async () => {
    vi.mocked(listPagesInSection).mockResolvedValueOnce([]);

    const { register } = await import('@/tools/listPages.js');
    await captureHandler(register)({ sectionId: 'sec-1' });

    expect(listPagesInSection).toHaveBeenCalledWith('sec-1', {
      limit: 50,
      orderBy: 'lastModified',
      order: 'desc',
    });
  });

  it('merges sections of a notebook, newest first, and caps the result', async () => {
    vi.mocked(collectSectionsForNotebook).mockResolvedValueOnce([
      sectionNode('sec-a', 'Section A'),
      sectionNode('sec-b', 'Section B', ['Group One']),
    ]);
    vi.mocked(listPagesInSection).mockImplementation(async (sectionId: string) =>
      sectionId === 'sec-a'
        ? [stubPage('a-old', 'A old', { lastModifiedDateTime: '2020-01-01T00:00:00Z' })]
        : [stubPage('b-new', 'B new', { lastModifiedDateTime: '2025-01-01T00:00:00Z' })],
    );

    const { register } = await import('@/tools/listPages.js');
    const result = parseResult(
      await captureHandler(register)({ notebookId: 'nb-1', limit: 5 }),
    );

    expect(result.sectionsScanned).toBe(2);
    expect(result.returned).toBe(2);
    const pages = result.pages as Record<string, unknown>[];
    expect(pages.map((page) => page.id)).toEqual(['b-new', 'a-old']);
    expect(pages[0]).toMatchObject({ section: 'Section B', groupPath: ['Group One'] });
  });

  it('honours title ordering ascending across sections', async () => {
    vi.mocked(collectSectionsForNotebook).mockResolvedValueOnce([
      sectionNode('sec-a', 'A'),
      sectionNode('sec-b', 'B'),
    ]);
    vi.mocked(listPagesInSection).mockImplementation(async (sectionId: string) =>
      sectionId === 'sec-a' ? [stubPage('a', 'Zebra')] : [stubPage('b', 'Apple')],
    );

    const { register } = await import('@/tools/listPages.js');
    const result = parseResult(
      await captureHandler(register)({ notebookId: 'nb-1', orderBy: 'title', order: 'asc' }),
    );

    expect((result.pages as { title: string }[]).map((page) => page.title)).toEqual([
      'Apple',
      'Zebra',
    ]);
  });

  it('refuses to guess a scope', async () => {
    const { register } = await import('@/tools/listPages.js');
    await expect(captureHandler(register)({})).rejects.toThrow('Provide either sectionId or notebookId');
  });
});

describe('find_pages handler', () => {
  it('forwards the query and spreads the result', async () => {
    vi.mocked(findPages).mockResolvedValueOnce({
      matches: [{ id: 'p1', title: 'Hit', matchedIn: 'title' }],
      scanned: { notebooks: 1, sections: 2, pages: 10, contentsFetched: 0 },
      contentScanTruncated: false,
      notes: [],
    });

    const { register } = await import('@/tools/findPages.js');
    const result = parseResult(
      await captureHandler(register)({ query: 'hit', notebookId: 'nb-1', limit: 5 }),
    );

    expect(findPages).toHaveBeenCalledWith({
      query: 'hit',
      sectionId: undefined,
      notebookId: 'nb-1',
      limit: 5,
      includeContent: undefined,
      maxContentPages: undefined,
    });
    expect(result.query).toBe('hit');
    expect(result.matches).toHaveLength(1);
  });
});

describe('get_notebook_tree handler', () => {
  it('summarises the hierarchy without page counts by default', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      {
        id: 'nb-1',
        name: 'Notebook One',
        isDefault: true,
        isShared: false,
        lastModified: '2024-06-01T00:00:00Z',
        webUrl: 'https://onenote.com/nb-1',
        sections: [sectionNode('sec-1', 'Notes'), sectionNode('sec-2', 'Ideas', ['Group One'])],
      },
    ] as never);

    const { register } = await import('@/tools/getNotebookTree.js');
    const result = parseResult(await captureHandler(register)({}));

    expect(countPagesForSections).not.toHaveBeenCalled();
    expect(result).toMatchObject({ notebookCount: 1, sectionCount: 2, pageCountsIncluded: false });
    const notebook = (result.notebooks as Record<string, unknown>[])[0]!;
    expect(notebook.sectionCount).toBe(2);
    expect((notebook.sections as Record<string, unknown>[])[1]).toEqual({
      id: 'sec-2',
      name: 'Ideas',
      groupPath: ['Group One'],
      lastModified: '2024-06-01T00:00:00Z',
    });
  });

  it('attaches page counts when requested', async () => {
    const sections = [sectionNode('sec-1', 'Notes')];
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([
      {
        id: 'nb-1',
        name: 'Notebook One',
        isDefault: true,
        isShared: false,
        lastModified: '2024-06-01T00:00:00Z',
        sections,
      },
    ] as never);
    vi.mocked(countPagesForSections).mockResolvedValueOnce([
      { ...sections[0]!, pageCount: 7 },
    ]);

    const { register } = await import('@/tools/getNotebookTree.js');
    const result = parseResult(await captureHandler(register)({ includePageCounts: true }));

    expect(countPagesForSections).toHaveBeenCalledWith(sections);
    expect(result.pageCountsIncluded).toBe(true);
    const notebook = (result.notebooks as Record<string, unknown>[])[0]!;
    expect((notebook.sections as Record<string, unknown>[])[0]!.pageCount).toBe(7);
  });

  it('passes a notebook id through to scope the tree', async () => {
    vi.mocked(collectNotebookTree).mockResolvedValueOnce([] as never);

    const { register } = await import('@/tools/getNotebookTree.js');
    await captureHandler(register)({ notebookId: 'nb-9' });

    expect(collectNotebookTree).toHaveBeenCalledWith('nb-9');
  });
});

describe('copy_page handler', () => {
  it('resolves both ends and reports the async copy', async () => {
    vi.mocked(getPage).mockResolvedValueOnce(stubPage('p1', 'Source Page'));
    vi.mocked(getSection).mockResolvedValueOnce(
      stubSection({ id: 'sec-2', displayName: 'Target', parentNotebook: { id: 'nb-2', displayName: 'Other Notebook' } }),
    );
    vi.mocked(copyPageToSection).mockResolvedValueOnce({
      status: 202,
      operationUrl: 'https://graph.microsoft.com/v1.0/operations/op-1',
    });

    const { register } = await import('@/tools/copyPage.js');
    const result = parseResult(
      await captureHandler(register)({ pageId: 'p1', targetSectionId: 'sec-2' }),
    );

    expect(result).toMatchObject({
      copied: true,
      status: 202,
      operationUrl: 'https://graph.microsoft.com/v1.0/operations/op-1',
      page: { id: 'p1', title: 'Source Page' },
      into: { id: 'sec-2', name: 'Target', notebook: 'Other Notebook' },
    });
    expect(String(result.note)).toContain('asynchronously');
  });

  it('resolves the section before issuing the copy so a bad id fails early', async () => {
    vi.mocked(getPage).mockResolvedValueOnce(stubPage('p1', 'Source Page'));
    vi.mocked(getSection).mockRejectedValueOnce(new Error('Section not found: nope'));

    const { register } = await import('@/tools/copyPage.js');
    await expect(
      captureHandler(register)({ pageId: 'p1', targetSectionId: 'nope' }),
    ).rejects.toThrow('Section not found');

    expect(copyPageToSection).not.toHaveBeenCalled();
  });
});

describe('auth_status handler', () => {
  it('reports the signed-in account and the token cache location', async () => {
    vi.mocked(getAuthSnapshot).mockResolvedValueOnce({
      signedIn: true,
      expiresOn: new Date('2026-09-20T03:02:53Z'),
      account: {
        username: 'someone@outlook.com',
        name: 'Someone',
        tenantId: 'tenant-1',
        environment: 'login.windows.net',
        homeAccountId: 'home-1',
        localAccountId: 'local-1',
      },
    });

    const { register } = await import('@/tools/authStatus.js');
    const result = parseResult(await captureHandler(register)({}));

    expect(result).toMatchObject({
      signedIn: true,
      clientId: 'test-client-id',
      accessTokenExpiresOn: '2026-09-20T03:02:53.000Z',
    });
    expect((result.account as Record<string, unknown>).username).toBe('someone@outlook.com');
    const cache = result.tokenCache as Record<string, unknown>;
    expect(String(cache.path)).toBe(join(tempRoot, 'onenote-mcp', 'tokens.json'));
    expect(cache.exists).toBe(false);
    expect(cache.configDirOverride).toBeNull();
    expect(result.remediation).toBeUndefined();
  });

  it('reports an existing cache file', async () => {
    vi.mocked(getAuthSnapshot).mockResolvedValueOnce({
      signedIn: true,
      expiresOn: undefined,
      account: undefined,
    });
    const dir = join(tempRoot, 'onenote-mcp');
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, 'tokens.json'), '{}');

    const { register } = await import('@/tools/authStatus.js');
    const result = parseResult(await captureHandler(register)({}));

    expect((result.tokenCache as Record<string, unknown>).exists).toBe(true);
  });

  it('offers remediation when the cache cannot be refreshed', async () => {
    vi.mocked(getAuthSnapshot).mockResolvedValueOnce({
      signedIn: false,
      expiresOn: undefined,
      account: undefined,
    });

    const { register } = await import('@/tools/authStatus.js');
    const result = parseResult(await captureHandler(register)({}));

    expect(result.signedIn).toBe(false);
    expect(result.account).toBeNull();
    expect(String(result.remediation)).toContain('login');
  });

  it('reflects a CONFIG_DIR override', async () => {
    process.env.ONENOTE_MCP_CONFIG_DIR = 'onenote-plus-mcp';
    vi.mocked(getAuthSnapshot).mockResolvedValueOnce({
      signedIn: true,
      expiresOn: undefined,
      account: undefined,
    });

    const { register } = await import('@/tools/authStatus.js');
    const result = parseResult(await captureHandler(register)({}));
    const cache = result.tokenCache as Record<string, unknown>;

    expect(String(cache.path)).toBe(join(tempRoot, 'onenote-plus-mcp', 'tokens.json'));
    expect(cache.configDirOverride).toBe('onenote-plus-mcp');
  });
});
