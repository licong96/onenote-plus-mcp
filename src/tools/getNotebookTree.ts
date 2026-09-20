import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  collectNotebookTree,
  countPagesForSections,
  type NotebookNode,
  type SectionNode,
} from '@/graph/tree.js';

const inputSchema = {
  notebookId: z
    .string()
    .optional()
    .describe('Notebook ID. Omit to map every notebook in the account.'),
  includePageCounts: z
    .boolean()
    .optional()
    .describe(
      'Include a page count per section. Costs one extra request per 100 pages per section, so it is slow on large accounts. Off by default.',
    ),
};

const shapeSection = (section: SectionNode): Record<string, unknown> => {
  const shaped: Record<string, unknown> = {
    id: section.id,
    name: section.name,
    groupPath: section.groupPath.length > 0 ? section.groupPath : undefined,
    lastModified: section.lastModified,
  };
  if (section.isDefault) shaped.isDefault = true;
  if (section.pageCount !== undefined) shaped.pageCount = section.pageCount;
  return shaped;
};

const shapeNotebook = (notebook: NotebookNode): Record<string, unknown> => ({
  id: notebook.id,
  name: notebook.name,
  isDefault: notebook.isDefault,
  isShared: notebook.isShared,
  lastModified: notebook.lastModified,
  webUrl: notebook.webUrl,
  sectionCount: notebook.sections.length,
  sections: notebook.sections.map(shapeSection),
});

export const register = (server: McpServer): void => {
  server.registerTool(
    'get_notebook_tree',
    {
      title: 'Get Notebook Tree',
      description:
        'Maps the OneNote structure: notebooks, the section groups inside them, and every section with its ID. Start here when you need section IDs for list_pages, create_page, or find_pages.',
      inputSchema,
    },
    async ({ notebookId, includePageCounts }) => {
      const notebooks = await collectNotebookTree(notebookId);

      const withCounts = includePageCounts
        ? await Promise.all(
            notebooks.map(async (notebook) => ({
              ...notebook,
              sections: await countPagesForSections(notebook.sections),
            })),
          )
        : notebooks;

      const totalSections = withCounts.reduce((sum, notebook) => sum + notebook.sections.length, 0);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                notebookCount: withCounts.length,
                sectionCount: totalSections,
                pageCountsIncluded: Boolean(includePageCounts),
                notebooks: withCounts.map(shapeNotebook),
              },
              null,
              2,
            ),
          },
        ],
      };
    },
  );
};
