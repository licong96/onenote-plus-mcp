import { beforeEach, describe, expect, it, vi } from 'vitest';
import { INDEX_VERSION, type PageIndex } from '@/index-store/store.js';

vi.mock('../../src/index-store/store.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@/index-store/store.js')>();
  return { ...actual, readIndex: vi.fn(), writeIndex: vi.fn() };
});

vi.mock('../../src/index-store/sync.js', () => ({
  syncIndex: vi.fn(),
}));

import { readIndex, writeIndex } from '@/index-store/store.js';
import { syncIndex } from '@/index-store/sync.js';
import {
  afterPageCopied,
  afterPageRemoved,
  afterPageContentChanged,
  afterSectionChanged,
} from '@/index-store/maintain.js';

const indexWith = (pageIds: string[]): PageIndex => ({
  version: INDEX_VERSION,
  generatedAt: '2024-06-01T00:00:00Z',
  notebooks: [
    {
      id: 'nb-1',
      name: 'NB',
      isDefault: false,
      isShared: false,
      lastModifiedDateTime: 'x',
      sections: [{ id: 'sec-1', name: 'A', groupPath: [], lastModifiedDateTime: 'x' }],
    },
  ],
  pages: pageIds.map((id) => ({
    id,
    title: id,
    notebookId: 'nb-1',
    notebookName: 'NB',
    sectionId: 'sec-1',
    sectionName: 'A',
    groupPath: [],
    createdDateTime: 'x',
    lastModifiedDateTime: 'x',
  })),
});

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(writeIndex).mockResolvedValue(undefined);
});

describe('when no index exists', () => {
  it('reports skipped rather than failing the write', async () => {
    vi.mocked(readIndex).mockResolvedValue(undefined);

    await expect(afterSectionChanged('sec-1')).resolves.toEqual({ status: 'skipped' });
    await expect(afterPageRemoved('p1')).resolves.toEqual({ status: 'skipped' });
  });
});

describe('afterSectionChanged', () => {
  it('syncs that section and persists', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith(['p1']));
    vi.mocked(syncIndex).mockResolvedValue(indexWith(['p1', 'p2']));

    const result = await afterSectionChanged('sec-1');

    expect(syncIndex).toHaveBeenCalledWith(expect.anything(), { sections: ['sec-1'] });
    expect(writeIndex).toHaveBeenCalled();
    expect(result).toEqual({ status: 'ok', sections: 1, pages: 2 });
  });

  it('reports failure without throwing when the refresh breaks', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));
    vi.mocked(syncIndex).mockRejectedValue(new Error('throttled'));

    const result = await afterSectionChanged('sec-1');

    expect(result.status).toBe('failed');
    expect(result.error).toContain('throttled');
  });
});

describe('afterPageRemoved', () => {
  it('drops the page without any network call', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith(['p1', 'p2']));

    const result = await afterPageRemoved('p1');

    expect(syncIndex).not.toHaveBeenCalled();
    const written = vi.mocked(writeIndex).mock.calls[0]![0];
    expect(written.pages.map((page) => page.id)).toEqual(['p2']);
    expect(result).toEqual({ status: 'ok', pages: 1 });
  });

  it('skips when the page was not indexed', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith(['p1']));

    const result = await afterPageRemoved('unknown');

    expect(writeIndex).not.toHaveBeenCalled();
    expect(result).toEqual({ status: 'skipped' });
  });

  it('reports failure without throwing when the write breaks', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith(['p1']));
    vi.mocked(writeIndex).mockRejectedValue(new Error('disk full'));

    const result = await afterPageRemoved('p1');

    expect(result.status).toBe('failed');
  });
});

describe('afterPageCopied', () => {
  it('refreshes the target section', async () => {
    vi.mocked(readIndex).mockResolvedValue(indexWith([]));
    vi.mocked(syncIndex).mockResolvedValue(indexWith(['new']));

    const result = await afterPageCopied('sec-1');

    expect(syncIndex).toHaveBeenCalledWith(expect.anything(), { sections: ['sec-1'] });
    expect(result.status).toBe('ok');
  });
});

describe('afterPageContentChanged', () => {
  it('is a no-op because page metadata did not change', () => {
    expect(afterPageContentChanged()).toEqual({ status: 'skipped' });
  });
});
