import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { listPagesInSection, type PageOrderBy, type SortOrder } from '@/graph/pages.js';
import { collectSectionsForNotebook } from '@/graph/tree.js';
import { DEFAULT_CONCURRENCY, mapWithConcurrency } from '@/util/concurrency.js';
import type { Page } from '@/graph/types.js';

const inputSchema = {
  sectionId: z
    .string()
    .optional()
    .describe('Section ID to list pages from. Provide this or notebookId.'),
  notebookId: z
    .string()
    .optional()
    .describe(
      'Notebook ID; lists pages across every section in that notebook (including sections nested in section groups). Provide this or sectionId.',
    ),
  limit: z
    .number()
    .int()
    .min(1)
    .max(500)
    .optional()
    .describe('Maximum pages to return (default 50, max 500).'),
  orderBy: z
    .enum(['title', 'lastModified', 'created'])
    .optional()
    .describe('Sort field (default lastModified).'),
  order: z
    .enum(['asc', 'desc'])
    .optional()
    .describe('Sort direction (default desc — newest or Z→A first).'),
};

interface Row {
  id: string;
  title: string;
  section?: string;
  notebook?: string;
  groupPath?: string[];
  created?: string;
  lastModified?: string;
  webUrl?: string;
}

const toRow = (page: Page, section?: string, notebook?: string, groupPath?: string[]): Row => ({
  id: page.id,
  title: page.title,
  section,
  notebook,
  groupPath: groupPath && groupPath.length > 0 ? groupPath : undefined,
  created: page.createdDateTime,
  lastModified: page.lastModifiedDateTime,
  webUrl: page.links?.oneNoteWebUrl?.href,
});

const sortKey = (row: Row, orderBy: PageOrderBy): string => {
  switch (orderBy) {
    case 'title':
      return row.title ?? '';
    case 'created':
      return row.created ?? '';
    default:
      return row.lastModified ?? '';
  }
};

const compareRows = (a: Row, b: Row, orderBy: PageOrderBy, order: SortOrder): number => {
  const result = sortKey(a, orderBy).localeCompare(sortKey(b, orderBy));
  return order === 'asc' ? result : -result;
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'list_pages',
    {
      title: 'List Pages',
      description:
        'Lists pages in a OneNote section, or across every section of a notebook, without needing a search term. Use it to browse what exists and get page IDs. Note: there is no account-wide page listing — Graph rejects it on accounts with many sections, so scope by section or notebook.',
      inputSchema,
    },
    async ({ sectionId, notebookId, limit, orderBy, order }) => {
      const effectiveLimit = limit ?? 50;
      const effectiveOrderBy: PageOrderBy = orderBy ?? 'lastModified';
      const effectiveOrder: SortOrder = order ?? 'desc';

      if (sectionId) {
        const pages = await listPagesInSection(sectionId, {
          limit: effectiveLimit,
          orderBy: effectiveOrderBy,
          order: effectiveOrder,
        });
        const rows = pages.map((page) => toRow(page, undefined, undefined, undefined));
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                { scope: { sectionId }, orderBy: effectiveOrderBy, order: effectiveOrder, returned: rows.length, pages: rows },
                null,
                2,
              ),
            },
          ],
        };
      }

      if (notebookId) {
        const sections = await collectSectionsForNotebook(notebookId);
        // Each section is asked for its own top `limit` — any page that belongs
        // in the notebook-wide top N must be in its section's top N, so this is
        // exact without fetching whole sections.
        const perSection = await mapWithConcurrency(
          sections,
          DEFAULT_CONCURRENCY,
          async (section) => ({
            section,
            pages: await listPagesInSection(section.id, {
              limit: effectiveLimit,
              orderBy: effectiveOrderBy,
              order: effectiveOrder,
            }),
          }),
        );

        const rows = perSection
          .flatMap(({ section, pages }) =>
            pages.map((page) => toRow(page, section.name, undefined, section.groupPath)),
          )
          .sort((a, b) => compareRows(a, b, effectiveOrderBy, effectiveOrder))
          .slice(0, effectiveLimit);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  scope: { notebookId },
                  sectionsScanned: sections.length,
                  orderBy: effectiveOrderBy,
                  order: effectiveOrder,
                  returned: rows.length,
                  pages: rows,
                },
                null,
                2,
              ),
            },
          ],
        };
      }

      throw new Error(
        'Provide either sectionId or notebookId. For an account-wide lookup use find_pages, and use get_notebook_tree to discover IDs.',
      );
    },
  );
};
