import { readIndex, writeIndex, type PageIndex } from './store.js';
import { syncIndex } from './sync.js';

/**
 * Keep the on-disk mirror truthful after a write.
 *
 * Every write path calls one of these so the index never silently disagrees
 * with OneNote. All of them are best-effort: a failed refresh is reported to
 * the caller but never fails the write itself, because the mutation has already
 * happened on Microsoft's side and rolling it back is not an option.
 *
 * These are *not* skipped when the index is absent — a missing index simply
 * means there is nothing to keep in sync, which is reported as `skipped`.
 */

export interface IndexUpdate {
  /** skipped when no index file exists; ok when refreshed; failed on error. */
  status: 'skipped' | 'ok' | 'failed';
  sections?: number;
  pages?: number;
  error?: string;
}

const refresh = async (
  index: PageIndex,
  sectionIds: readonly string[],
): Promise<IndexUpdate> => {
  try {
    const next = await syncIndex(index, { sections: [...sectionIds] });
    await writeIndex(next);
    return { status: 'ok', sections: sectionIds.length, pages: next.pages.length };
  } catch (error) {
    return { status: 'failed', error: (error as Error).message };
  }
};

/** Call after a page is created in, or moved into, `sectionId`. */
export const afterSectionChanged = async (sectionId: string): Promise<IndexUpdate> => {
  const index = await readIndex();
  if (!index) return { status: 'skipped' };
  return refresh(index, [sectionId]);
};

/** Call after a page is deleted, to drop it without a network round trip. */
export const afterPageRemoved = async (pageId: string): Promise<IndexUpdate> => {
  const index = await readIndex();
  if (!index) return { status: 'skipped' };

  const remaining = index.pages.filter((page) => page.id !== pageId);
  if (remaining.length === index.pages.length) return { status: 'skipped' };

  try {
    await writeIndex({ ...index, pages: remaining });
    return { status: 'ok', pages: remaining.length };
  } catch (error) {
    return { status: 'failed', error: (error as Error).message };
  }
};

/**
 * Call after a page is copied: the source section is untouched, but the target
 * section gained a page, and the new page carries a brand-new ID.
 */
export const afterPageCopied = async (targetSectionId: string): Promise<IndexUpdate> =>
  afterSectionChanged(targetSectionId);

/** Call after a page's content changes. Metadata is unchanged, so this is a no-op. */
export const afterPageContentChanged = (): IndexUpdate => ({ status: 'skipped' });
