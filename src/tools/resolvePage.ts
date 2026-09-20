import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { resolvePage } from '@/index-store/resolve.js';

const inputSchema = {
  pageId: z.string().optional().describe('Exact page ID, if you already have it.'),
  title: z.string().optional().describe('Page title as shown in OneNote.'),
  notebook: z.string().optional().describe('Notebook name, to disambiguate duplicate titles.'),
  section: z.string().optional().describe('Section name, to disambiguate duplicate titles.'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'resolve_page',
    {
      title: 'Resolve Page',
      description:
        'Turns a human description of a page ("今日任务 / 香港", or a title alone) into a real page ID, served from the local index. Use this before reading or editing a page you named rather than identified. Stored IDs go stale when a page is moved, so this verifies against Graph on a miss and repairs the index.',
      inputSchema,
    },
    async ({ pageId, title, notebook, section }) => {
      if (!pageId && !title) {
        throw new Error('Provide pageId, or title (optionally with notebook and section).');
      }

      const result = await resolvePage({ pageId, title, notebook, section });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                resolved: Boolean(result.page),
                source: result.source,
                indexRepaired: result.repaired,
                page: result.page
                  ? {
                      id: result.page.id,
                      title: result.page.title,
                      notebook: result.page.notebookName,
                      section: result.page.sectionName,
                      lastModified: result.page.lastModifiedDateTime,
                      webUrl: result.page.webUrl,
                    }
                  : null,
                candidates: result.candidates.map((page) => ({
                  id: page.id,
                  title: page.title,
                  notebook: page.notebookName,
                  section: page.sectionName,
                  lastModified: page.lastModifiedDateTime,
                })),
                notes: result.notes,
                indexStats: result.stats,
                guidance:
                  result.source === 'none'
                    ? 'Not found. Try index mode "search" for a fuzzy match, or "sync" if pages changed recently.'
                    : undefined,
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
