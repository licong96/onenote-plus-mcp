import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { copyPageToSection, listAllPagesInSection, listPagesInSection } from '@/graph/pages.js';

vi.mock('../../src/auth/index.js', () => ({
  getAccessToken: vi.fn(async () => 'fake-access-token'),
}));

const fetchMock = vi.fn();
const originalFetch = globalThis.fetch;

const jsonResponse = (body: unknown, init: ResponseInit = {}): Response =>
  new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    ...init,
  });

const pageRow = (id: string): Record<string, string> => ({
  id,
  title: `Page ${id}`,
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-06-01T00:00:00Z',
  contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
});

const urlOf = (callIndex = 0): URL => new URL(String(fetchMock.mock.calls[callIndex]![0]));

beforeEach(() => {
  fetchMock.mockReset();
  globalThis.fetch = fetchMock as typeof fetch;
});

afterEach(() => {
  globalThis.fetch = originalFetch;
});

describe('listPagesInSection', () => {
  it('orders server-side newest-first and requests a page-size hint', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [pageRow('p1')] }));

    await listPagesInSection('sec-1', { limit: 5 });

    const url = urlOf();
    expect(url.pathname).toBe('/v1.0/me/onenote/sections/sec-1/pages');
    expect(url.searchParams.get('$orderby')).toBe('lastModifiedDateTime desc');
    expect(url.searchParams.get('$top')).toBe('5');
    expect(url.searchParams.get('$select')).toContain('title');
  });

  it('defaults to 50 rows, newest first', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));

    await listPagesInSection('sec-1');

    const url = urlOf();
    expect(url.searchParams.get('$orderby')).toBe('lastModifiedDateTime desc');
    expect(url.searchParams.get('$top')).toBe('50');
  });

  it('sorts by title or creation date when asked', async () => {
    // Fresh Response per call: a body can only be consumed once, so reusing one
    // object across fetches fails with "Body has already been read".
    fetchMock.mockImplementation(async () => jsonResponse({ value: [] }));

    await listPagesInSection('sec-1', { orderBy: 'title', order: 'asc' });
    expect(urlOf(0).searchParams.get('$orderby')).toBe('title asc');

    await listPagesInSection('sec-1', { orderBy: 'created', order: 'asc' });
    expect(urlOf(1).searchParams.get('$orderby')).toBe('createdDateTime asc');
  });

  it('stops paginating once the limit is satisfied', async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse({
        value: [pageRow('p1'), pageRow('p2')],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/next-page',
      }),
    );

    const pages = await listPagesInSection('sec-1', { limit: 2 });

    expect(pages.map((page) => page.id)).toEqual(['p1', 'p2']);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('follows nextLink when the first page is short of the limit', async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: [pageRow('p1')],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/next-page',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [pageRow('p2')] }));

    const pages = await listPagesInSection('sec-1', { limit: 2 });

    expect(pages.map((page) => page.id)).toEqual(['p1', 'p2']);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('caps the per-request page size at 100 while still collecting the full limit', async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: [pageRow('p1')],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/next-page',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [pageRow('p2')] }));

    await listPagesInSection('sec-1', { limit: 250 });

    expect(urlOf(0).searchParams.get('$top')).toBe('100');
  });

  it('encodes section IDs containing "!"', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));
    await listPagesInSection('0-B18ADB61AADA4FE8!267');
    expect(urlOf().pathname).toContain('0-B18ADB61AADA4FE8!267');
  });
});

describe('listAllPagesInSection', () => {
  it('walks every page without ordering or an artificial cap', async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: [pageRow('p1')],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/next-page',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [pageRow('p2')] }));

    const pages = await listAllPagesInSection('sec-1');

    expect(pages.map((page) => page.id)).toEqual(['p1', 'p2']);
    const url = urlOf();
    expect(url.searchParams.get('$top')).toBe('100');
    expect(url.searchParams.has('$orderby')).toBe(false);
  });
});

describe('copyPageToSection', () => {
  it('POSTs the target section id and surfaces the async job URL', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response('', {
        status: 202,
        headers: { 'Operation-Location': 'https://graph.microsoft.com/v1.0/operations/op-1' },
      }),
    );

    const result = await copyPageToSection('page-1', 'sec-2');

    expect(result).toEqual({
      status: 202,
      operationUrl: 'https://graph.microsoft.com/v1.0/operations/op-1',
    });

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(String(url)).toBe(
      'https://graph.microsoft.com/v1.0/me/onenote/pages/page-1/copyToSection',
    );
    expect((init as RequestInit).method).toBe('POST');
    expect(JSON.parse(String((init as RequestInit).body))).toEqual({ id: 'sec-2' });
  });

  it('tolerates a 202 with no operation URL', async () => {
    fetchMock.mockResolvedValueOnce(new Response('', { status: 202 }));

    const result = await copyPageToSection('page-1', 'sec-2');
    expect(result.status).toBe(202);
    expect(result.operationUrl).toBeUndefined();
  });

  it('throws a shaped GraphError when the target section is rejected', async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse(
        { error: { code: 'NotFound', message: 'Section not found' } },
        { status: 404, statusText: 'Not Found' },
      ),
    );

    await expect(copyPageToSection('page-1', 'missing')).rejects.toThrow('Section not found');
  });
});
