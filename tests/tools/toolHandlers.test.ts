import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Notebook, Section, SectionGroup, Page } from '@/graph/types.js';

vi.mock('../../src/graph/notebooks.js', () => ({
  listNotebooks: vi.fn(),
  createNotebook: vi.fn(),
}));

vi.mock('../../src/graph/sections.js', () => ({
  listSections: vi.fn(),
  createSection: vi.fn(),
}));

vi.mock('../../src/graph/sectionGroups.js', () => ({
  listSectionGroups: vi.fn(),
  createSectionGroup: vi.fn(),
}));

vi.mock('../../src/graph/pages.js', () => ({
  searchPages: vi.fn(),
  getPage: vi.fn(),
  getPageContent: vi.fn(),
  createPage: vi.fn(),
  deletePage: vi.fn(),
  updatePage: vi.fn(),
}));

vi.mock('../../src/markdown.js', () => ({
  markdownToHtmlFragment: vi.fn((md: string) => `<p>${md}</p>`),
  htmlToMarkdown: vi.fn((html: string) => `markdown:${html}`),
  markdownToOneNoteHtml: vi.fn((_md: string, title: string) => `<html><head><title>${title}</title></head><body>converted</body></html>`),
  htmlToOneNotePage: vi.fn((html: string, title: string) => `<html><head><title>${title}</title></head><body>${html}</body></html>`),
}));

import { listNotebooks, createNotebook } from '@/graph/notebooks.js';
import { listSections, createSection } from '@/graph/sections.js';
import { listSectionGroups, createSectionGroup } from '@/graph/sectionGroups.js';
import { searchPages, getPage, getPageContent, createPage, deletePage, updatePage } from '@/graph/pages.js';
import { markdownToHtmlFragment } from '@/markdown.js';

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

const parseResult = (result: ToolResult): unknown =>
  JSON.parse(result.content[0]!.text);

const stubNotebook = (overrides: Partial<Notebook> = {}): Notebook => ({
  id: 'nb-1',
  displayName: 'Test Notebook',
  isDefault: true,
  isShared: false,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  links: {
    oneNoteWebUrl: { href: 'https://onenote.com/nb-1' },
    oneNoteClientUrl: { href: 'onenote:nb-1' },
  },
  ...overrides,
});

const stubSection = (overrides: Partial<Section> = {}): Section => ({
  id: 'sec-1',
  displayName: 'Test Section',
  isDefault: false,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  parentNotebook: { id: 'nb-1', displayName: 'Test Notebook' },
  parentSectionGroup: null,
  links: {
    oneNoteWebUrl: { href: 'https://onenote.com/sec-1' },
    oneNoteClientUrl: { href: 'onenote:sec-1' },
  },
  ...overrides,
});

const stubSectionGroup = (overrides: Partial<SectionGroup> = {}): SectionGroup => ({
  id: 'sg-1',
  displayName: 'Test Group',
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  parentNotebook: { id: 'nb-1', displayName: 'Test Notebook' },
  parentSectionGroup: null,
  links: {
    oneNoteWebUrl: { href: 'https://onenote.com/sg-1' },
    oneNoteClientUrl: { href: 'onenote:sg-1' },
  },
  ...overrides,
});

const stubPage = (overrides: Partial<Page> = {}): Page => ({
  id: 'page-1',
  title: 'Test Page',
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-1/content',
  parentSection: { id: 'sec-1', displayName: 'Test Section' },
  parentNotebook: { id: 'nb-1', displayName: 'Test Notebook' },
  links: {
    oneNoteWebUrl: { href: 'https://onenote.com/page-1' },
    oneNoteClientUrl: { href: 'onenote:page-1' },
  },
  ...overrides,
});

beforeEach(() => {
  vi.clearAllMocks();
});

describe('list_notebooks handler', () => {
  it('maps notebook fields to a summary array', async () => {
    vi.mocked(listNotebooks).mockResolvedValueOnce([
      stubNotebook(),
      stubNotebook({ id: 'nb-2', displayName: 'Second', isDefault: false, isShared: true }),
    ]);

    const { register } = await import('@/tools/listNotebooks.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({})) as unknown[];

    expect(result).toHaveLength(2);
    expect(result[0]).toEqual({
      id: 'nb-1',
      name: 'Test Notebook',
      isDefault: true,
      isShared: false,
      webUrl: 'https://onenote.com/nb-1',
      lastModified: '2024-06-01T00:00:00Z',
    });
    expect((result[1] as Record<string, unknown>).isShared).toBe(true);
  });

  it('handles notebooks without links gracefully', async () => {
    vi.mocked(listNotebooks).mockResolvedValueOnce([
      stubNotebook({ links: undefined }),
    ]);

    const { register } = await import('@/tools/listNotebooks.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({})) as Record<string, unknown>[];

    expect(result[0]!.webUrl).toBeUndefined();
  });
});

describe('create_notebook handler', () => {
  it('returns the created notebook summary', async () => {
    vi.mocked(createNotebook).mockResolvedValueOnce(stubNotebook({ id: 'nb-new' }));

    const { register } = await import('@/tools/createNotebook.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ name: 'New Book' }));

    expect(createNotebook).toHaveBeenCalledWith('New Book');
    expect(result).toEqual({
      id: 'nb-new',
      name: 'Test Notebook',
      isDefault: true,
      webUrl: 'https://onenote.com/nb-1',
      createdDateTime: '2024-01-01T00:00:00Z',
    });
  });
});

describe('list_sections handler', () => {
  it('maps section fields with parent info', async () => {
    vi.mocked(listSections).mockResolvedValueOnce([stubSection()]);

    const { register } = await import('@/tools/listSections.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ notebookId: 'nb-1' })) as Record<string, unknown>[];

    expect(listSections).toHaveBeenCalledWith('nb-1');
    expect(result[0]).toEqual({
      id: 'sec-1',
      name: 'Test Section',
      isDefault: false,
      notebook: 'Test Notebook',
      notebookId: 'nb-1',
      webUrl: 'https://onenote.com/sec-1',
      lastModified: '2024-06-01T00:00:00Z',
    });
  });

  it('passes undefined notebookId to list all sections', async () => {
    vi.mocked(listSections).mockResolvedValueOnce([]);

    const { register } = await import('@/tools/listSections.js');
    const handler = captureHandler(register);
    await handler({ notebookId: undefined });

    expect(listSections).toHaveBeenCalledWith(undefined);
  });
});

describe('create_section handler', () => {
  it('creates a section under a notebook', async () => {
    vi.mocked(createSection).mockResolvedValueOnce(stubSection({ id: 'sec-new' }));

    const { register } = await import('@/tools/createSection.js');
    const handler = captureHandler(register);
    const result = parseResult(
      await handler({ notebookId: 'nb-1', sectionGroupId: undefined, name: 'Work' }),
    ) as Record<string, unknown>;

    expect(createSection).toHaveBeenCalledWith({ notebookId: 'nb-1', sectionGroupId: undefined }, 'Work');
    expect(result.id).toBe('sec-new');
    expect(result.notebook).toBe('Test Notebook');
  });

  it('creates a section under a section group', async () => {
    vi.mocked(createSection).mockResolvedValueOnce(
      stubSection({
        parentNotebook: undefined,
        parentSectionGroup: { id: 'sg-1', displayName: 'Group' },
      }),
    );

    const { register } = await import('@/tools/createSection.js');
    const handler = captureHandler(register);
    const result = parseResult(
      await handler({ notebookId: undefined, sectionGroupId: 'sg-1', name: 'Sub' }),
    ) as Record<string, unknown>;

    expect(result.sectionGroup).toBe('Group');
    expect(result.sectionGroupId).toBe('sg-1');
  });

  it('throws when both notebookId and sectionGroupId are provided', async () => {
    const { register } = await import('@/tools/createSection.js');
    const handler = captureHandler(register);

    await expect(
      handler({ notebookId: 'nb-1', sectionGroupId: 'sg-1', name: 'Bad' }),
    ).rejects.toThrow(/exactly one/i);
  });

  it('throws when neither notebookId nor sectionGroupId is provided', async () => {
    const { register } = await import('@/tools/createSection.js');
    const handler = captureHandler(register);

    await expect(
      handler({ notebookId: undefined, sectionGroupId: undefined, name: 'Bad' }),
    ).rejects.toThrow(/exactly one/i);
  });
});

describe('list_section_groups handler', () => {
  it('maps section group fields including parent info', async () => {
    vi.mocked(listSectionGroups).mockResolvedValueOnce([
      stubSectionGroup({
        parentSectionGroup: { id: 'sg-parent', displayName: 'Parent Group' },
      }),
    ]);

    const { register } = await import('@/tools/listSectionGroups.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ notebookId: 'nb-1' })) as Record<string, unknown>[];

    expect(result[0]).toEqual({
      id: 'sg-1',
      name: 'Test Group',
      notebook: 'Test Notebook',
      notebookId: 'nb-1',
      parentSectionGroup: 'Parent Group',
      parentSectionGroupId: 'sg-parent',
      webUrl: 'https://onenote.com/sg-1',
      lastModified: '2024-06-01T00:00:00Z',
    });
  });
});

describe('create_section_group handler', () => {
  it('creates under a notebook and returns formatted result', async () => {
    vi.mocked(createSectionGroup).mockResolvedValueOnce(stubSectionGroup({ id: 'sg-new' }));

    const { register } = await import('@/tools/createSectionGroup.js');
    const handler = captureHandler(register);
    const result = parseResult(
      await handler({ notebookId: 'nb-1', sectionGroupId: undefined, name: 'New Group' }),
    ) as Record<string, unknown>;

    expect(createSectionGroup).toHaveBeenCalledWith(
      { notebookId: 'nb-1', sectionGroupId: undefined },
      'New Group',
    );
    expect(result.id).toBe('sg-new');
    expect(result.createdDateTime).toBe('2024-01-01T00:00:00Z');
  });

  it('creates under a parent section group', async () => {
    vi.mocked(createSectionGroup).mockResolvedValueOnce(
      stubSectionGroup({
        parentNotebook: undefined,
        parentSectionGroup: { id: 'sg-parent', displayName: 'Parent' },
      }),
    );

    const { register } = await import('@/tools/createSectionGroup.js');
    const handler = captureHandler(register);
    const result = parseResult(
      await handler({ notebookId: undefined, sectionGroupId: 'sg-parent', name: 'Nested' }),
    ) as Record<string, unknown>;

    expect(result.parentSectionGroup).toBe('Parent');
    expect(result.parentSectionGroupId).toBe('sg-parent');
  });

  it('throws when both notebookId and sectionGroupId are provided', async () => {
    const { register } = await import('@/tools/createSectionGroup.js');
    const handler = captureHandler(register);

    await expect(
      handler({ notebookId: 'nb-1', sectionGroupId: 'sg-1', name: 'Bad' }),
    ).rejects.toThrow(/exactly one/i);
  });

  it('throws when neither parent is provided', async () => {
    const { register } = await import('@/tools/createSectionGroup.js');
    const handler = captureHandler(register);

    await expect(
      handler({ notebookId: undefined, sectionGroupId: undefined, name: 'Bad' }),
    ).rejects.toThrow(/exactly one/i);
  });
});

describe('search_pages handler', () => {
  it('maps page fields to a summary array', async () => {
    vi.mocked(searchPages).mockResolvedValueOnce([stubPage()]);

    const { register } = await import('@/tools/searchPages.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ query: 'test', limit: 10 })) as Record<string, unknown>[];

    expect(searchPages).toHaveBeenCalledWith({ query: 'test', limit: 10 });
    expect(result[0]).toEqual({
      id: 'page-1',
      title: 'Test Page',
      notebook: 'Test Notebook',
      section: 'Test Section',
      webUrl: 'https://onenote.com/page-1',
      lastModified: '2024-06-01T00:00:00Z',
    });
  });
});

describe('read_page handler', () => {
  it('returns HTML content by default', async () => {
    vi.mocked(getPage).mockResolvedValueOnce(stubPage());
    vi.mocked(getPageContent).mockResolvedValueOnce('<html><body>hello</body></html>');

    const { register } = await import('@/tools/readPage.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ pageId: 'page-1', format: undefined })) as Record<string, unknown>;

    expect(getPage).toHaveBeenCalledWith('page-1');
    expect(getPageContent).toHaveBeenCalledWith('page-1');
    expect(result.format).toBe('html');
    expect(result.content).toBe('<html><body>hello</body></html>');
    expect(result.title).toBe('Test Page');
    expect(result.notebook).toBe('Test Notebook');
    expect(result.section).toBe('Test Section');
    expect(result.links).toEqual({
      web: 'https://onenote.com/page-1',
      client: 'onenote:page-1',
    });
  });

  it('converts to markdown when format is "markdown"', async () => {
    vi.mocked(getPage).mockResolvedValueOnce(stubPage());
    vi.mocked(getPageContent).mockResolvedValueOnce('<p>hello</p>');

    const { register } = await import('@/tools/readPage.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ pageId: 'page-1', format: 'markdown' })) as Record<string, unknown>;

    expect(result.format).toBe('markdown');
    expect(result.content).toBe('markdown:<p>hello</p>');
  });
});

describe('delete_page handler', () => {
  it('deletes the page and returns a confirmation', async () => {
    vi.mocked(deletePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/deletePage.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({ pageId: 'page-42' }));

    expect(deletePage).toHaveBeenCalledWith('page-42');
    expect(result).toEqual({ deleted: true, pageId: 'page-42' });
  });
});

describe('update_page handler', () => {
  it('converts markdown content via markdownToHtmlFragment', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'body', action: 'append', content: 'hello world', format: 'markdown' }],
    });

    expect(markdownToHtmlFragment).toHaveBeenCalledWith('hello world');
    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]).toEqual({
      target: 'body',
      action: 'append',
      content: '<p>hello world</p>',
    });
  });

  it('passes HTML content through without conversion', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'body', action: 'append', content: '<b>bold</b>', format: 'html' }],
    });

    expect(markdownToHtmlFragment).not.toHaveBeenCalled();
    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('<b>bold</b>');
  });

  it('omits content for delete actions', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: '#data-id-123', action: 'delete', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]).toEqual({ target: '#data-id-123', action: 'delete' });
    expect(commands[0]).not.toHaveProperty('content');
  });

  it('includes position for insert actions', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{
        target: '#el-1',
        action: 'insert',
        position: 'after',
        content: 'inserted',
        format: 'markdown',
      }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.position).toBe('after');
    expect(commands[0]!.content).toBe('<p>inserted</p>');
  });

  it('unwraps <p> wrapping when targeting "title" with markdown', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'title', action: 'replace', content: 'New Title', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('New Title');
  });

  it('unwraps <p> wrapping when targeting "#title" (hash prefix)', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: '#title', action: 'replace', content: 'Hash Title', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('Hash Title');
  });

  it('does NOT unwrap <p> for title when format is html', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'title', action: 'replace', content: '<p>Keep Wrapped</p>', format: 'html' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('<p>Keep Wrapped</p>');
  });

  it('does NOT unwrap non-title targets', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'body', action: 'append', content: 'body text', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('<p>body text</p>');
  });

  it('greedy-unwraps outer <p> pair for multi-paragraph title markdown', async () => {
    vi.mocked(markdownToHtmlFragment).mockReturnValueOnce('<p>line1</p>\n<p>line2</p>');
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'title', action: 'replace', content: 'line1\n\nline2', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('line1</p>\n<p>line2');
  });

  it('does NOT unwrap when content has no <p> wrapper at all', async () => {
    vi.mocked(markdownToHtmlFragment).mockReturnValueOnce('<h1>heading</h1>');
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    await handler({
      pageId: 'page-1',
      operations: [{ target: 'title', action: 'replace', content: '# heading', format: 'markdown' }],
    });

    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands[0]!.content).toBe('<h1>heading</h1>');
  });

  it('maps multiple operations and returns the correct count', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({
      pageId: 'page-1',
      operations: [
        { target: 'body', action: 'append', content: 'first', format: 'markdown' },
        { target: '#el-2', action: 'delete', format: 'markdown' },
        { target: 'title', action: 'replace', content: 'New Title', format: 'markdown' },
      ],
    }));

    expect(updatePage).toHaveBeenCalledWith('page-1', expect.any(Array));
    const commands = vi.mocked(updatePage).mock.calls[0]![1];
    expect(commands).toHaveLength(3);
    expect(result).toEqual({ updated: true, pageId: 'page-1', operationCount: 3 });
  });

  it('returns success response with correct shape', async () => {
    vi.mocked(updatePage).mockResolvedValueOnce(undefined);

    const { register } = await import('@/tools/updatePage.js');
    const handler = captureHandler(register);
    const raw = await handler({
      pageId: 'page-99',
      operations: [{ target: 'body', action: 'append', content: 'x', format: 'html' }],
    });

    expect(raw.content).toHaveLength(1);
    expect(raw.content[0]!.type).toBe('text');
    const parsed = JSON.parse(raw.content[0]!.text);
    expect(parsed).toEqual({ updated: true, pageId: 'page-99', operationCount: 1 });
  });
});

describe('create_page handler', () => {
  it('creates a page with markdown format and returns summary', async () => {
    vi.mocked(createPage).mockResolvedValueOnce(stubPage({ id: 'page-new', title: 'Created' }));

    const { register } = await import('@/tools/createPage.js');
    const handler = captureHandler(register);
    const result = parseResult(await handler({
      sectionId: 'sec-1',
      title: 'Created',
      content: '# Hello',
      format: 'markdown',
      attachments: undefined,
    })) as Record<string, unknown>;

    expect(result.id).toBe('page-new');
    expect(result.title).toBe('Created');
    expect(result.attachmentCount).toBe(0);
  });

  it('creates a page with html format', async () => {
    vi.mocked(createPage).mockResolvedValueOnce(stubPage());

    const { register } = await import('@/tools/createPage.js');
    const handler = captureHandler(register);
    await handler({
      sectionId: 'sec-1',
      title: 'HTML Page',
      content: '<p>raw html</p>',
      format: 'html',
      attachments: undefined,
    });

    expect(createPage).toHaveBeenCalledWith(
      expect.objectContaining({ sectionId: 'sec-1' }),
    );
  });
});
