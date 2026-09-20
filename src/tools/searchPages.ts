import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { searchPages } from '@/graph/pages.js';

const inputSchema = {
  query: z.string().min(1).describe('Full-text search query against page titles and content.'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .optional()
    .describe('Maximum number of pages to return (default 25, max 100).'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'search_pages',
    {
      title: 'Search Pages',
      description:
        'Searches OneNote pages by full-text query. Returns matching pages with IDs, titles, and parent section/notebook info.',
      inputSchema,
    },
    async ({ query, limit }) => {
      const pages = await searchPages({ query, limit });
      const summary = pages.map((p) => ({
        id: p.id,
        title: p.title,
        notebook: p.parentNotebook?.displayName,
        section: p.parentSection?.displayName,
        webUrl: p.links?.oneNoteWebUrl?.href,
        lastModified: p.lastModifiedDateTime,
      }));
      return {
        content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }],
      };
    },
  );
};
