import { graphRequest, graphRequestRaw, paginate, paginateUntil } from './client.js';
import type { Page } from './types.js';

const PAGE_SELECT =
  'id,title,createdDateTime,lastModifiedDateTime,contentUrl,parentSection,parentNotebook,links';

const PAGE_EXPAND =
  'parentSection($select=id,displayName),parentNotebook($select=id,displayName)';

export interface SearchPagesOptions {
  query: string;
  limit?: number;
}

export const searchPages = async (options: SearchPagesOptions): Promise<Page[]> => {
  const { query, limit = 25 } = options;
  // Graph's $search on /me/onenote/pages does a server-side full-text match on
  // page title + content. Quoting the term keeps it safe across odata.
  const escaped = query.replaceAll('"', '\\"');
  const top = Math.max(1, Math.min(limit, 100));
  const pages = await paginate<Page>('/me/onenote/pages', {
    query: {
      $search: `"${escaped}"`,
      $select: PAGE_SELECT,
      $expand: PAGE_EXPAND,
      $top: top,
    },
  });
  // $top is a page size, not a total cap — paginate follows nextLinks until exhausted.
  return pages.slice(0, limit);
};

export const getPage = (pageId: string): Promise<Page> =>
  graphRequest<Page>(`/me/onenote/pages/${encodeURIComponent(pageId)}`, {
    query: {
      $select: PAGE_SELECT,
      $expand: PAGE_EXPAND,
    },
  });

export const getPageContent = (pageId: string): Promise<string> =>
  graphRequest<string>(`/me/onenote/pages/${encodeURIComponent(pageId)}/content`, {
    accept: 'text/html',
    parse: 'text',
  });

export interface CreatePageAttachment {
  /** The part name referenced from the HTML via `name:<name>` URIs. */
  name: string;
  contentType: string;
  /** Raw bytes for the part. */
  data: Uint8Array;
}

export interface CreatePageOptions {
  sectionId: string;
  /** Full HTML document including <html><head><title>…</title></head><body>…</body></html>. */
  html: string;
  /** Optional binary attachments. When present, the page is sent as multipart/form-data. */
  attachments?: CreatePageAttachment[];
}

export const createPage = (options: CreatePageOptions): Promise<Page> => {
  const { sectionId, html, attachments } = options;
  const path = `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`;

  if (!attachments || attachments.length === 0) {
    return graphRequest<Page>(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/xhtml+xml' },
      body: html,
    });
  }

  const form = new FormData();
  form.append('Presentation', new Blob([html], { type: 'application/xhtml+xml' }));
  for (const part of attachments) {
    form.append(part.name, new Blob([part.data], { type: part.contentType }));
  }
  // Don't set Content-Type — fetch + FormData adds the multipart boundary itself.
  return graphRequest<Page>(path, {
    method: 'POST',
    body: form,
  });
};

export const deletePage = async (pageId: string): Promise<void> => {
  await graphRequest<void>(`/me/onenote/pages/${encodeURIComponent(pageId)}`, {
    method: 'DELETE',
    parse: 'none',
  });
};

export type UpdatePageAction = 'append' | 'prepend' | 'insert' | 'replace' | 'delete';

export interface UpdatePageCommand {
  /** `body`, `title`, or the element's `data-id` value. A leading `#` is accepted and stripped. */
  target: string;
  action: UpdatePageAction;
  /** Required when `action` is `insert`; optional with `append`/`prepend` for sibling positioning. */
  position?: 'before' | 'after';
  /** HTML fragment. Required for every action except `delete`. */
  content?: string;
}

export const updatePage = async (
  pageId: string,
  commands: UpdatePageCommand[],
): Promise<void> => {
  // Graph expects raw data-id values, not CSS-style "#abc" selectors. Strip the
  // leading `#` if the caller passed one — they're easy to copy that way from
  // read_page HTML and we don't want them to 400 the request.
  const sanitized = commands.map((cmd) => ({
    ...cmd,
    target: cmd.target.startsWith('#') ? cmd.target.slice(1) : cmd.target,
  }));
  await graphRequest<void>(`/me/onenote/pages/${encodeURIComponent(pageId)}/content`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(sanitized),
    parse: 'none',
  });
};

/**
 * Metadata-only page select for listing.
 *
 * Parent notebook/section are deliberately omitted: with a section-scoped query
 * the caller already knows both, and dropping them lets us add `$orderby`
 * without betting on `$expand` + `$orderby` being accepted together.
 */
const PAGE_LIST_SELECT = 'id,title,createdDateTime,lastModifiedDateTime,contentUrl,links';

export type PageOrderBy = 'title' | 'lastModified' | 'created';

export type SortOrder = 'asc' | 'desc';

export interface ListPagesOptions {
  limit?: number;
  orderBy?: PageOrderBy;
  order?: SortOrder;
}

const orderByField = (orderBy: PageOrderBy): string => {
  switch (orderBy) {
    case 'title':
      return 'title';
    case 'created':
      return 'createdDateTime';
    default:
      return 'lastModifiedDateTime';
  }
};

/**
 * Pages of one section, newest-first by default.
 *
 * Ordering is done server-side and pagination stops as soon as `limit` rows are
 * collected, so asking for the 10 most recent pages of a 500-page section costs
 * one request rather than eleven.
 */
export const listPagesInSection = (
  sectionId: string,
  options: ListPagesOptions = {},
): Promise<Page[]> => {
  const { limit = 50, orderBy = 'lastModified', order = 'desc' } = options;
  return paginateUntil<Page>(
    `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`,
    limit,
    {
      query: {
        $select: PAGE_LIST_SELECT,
        $orderby: `${orderByField(orderBy)} ${order}`,
        $top: Math.min(limit, 100),
      },
    },
  );
};

/** Every page of one section (metadata only) — used by whole-section scans. */
export const listAllPagesInSection = (sectionId: string): Promise<Page[]> =>
  paginate<Page>(`/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`, {
    query: { $select: PAGE_LIST_SELECT, $top: 100 },
  });

export interface CopyPageResult {
  /** 202 for an accepted async copy; Graph completes the job shortly after. */
  status: number;
  /** `Operation-Location` header, when Graph returns one, for polling the job. */
  operationUrl?: string;
}

/**
 * Copy a page into another section (Graph `copyToSection`).
 *
 * This is a copy, never a move — the source page is untouched. Moving is
 * intentionally not offered: Graph has no move endpoint, so a move is
 * copy-then-delete, and that delete is irreversible.
 */
export const copyPageToSection = async (
  pageId: string,
  targetSectionId: string,
): Promise<CopyPageResult> => {
  const { status, headers } = await graphRequestRaw(
    `/me/onenote/pages/${encodeURIComponent(pageId)}/copyToSection`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: targetSectionId }),
    },
  );
  // Headers.get is case-insensitive, so one lookup covers both spellings.
  return { status, operationUrl: headers.get('operation-location') ?? undefined };
};
