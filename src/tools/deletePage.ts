import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { deletePage } from '@/graph/pages.js';

const inputSchema = {
  pageId: z.string().min(1).describe('OneNote page ID to delete. This action is irreversible.'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'delete_page',
    {
      title: 'Delete Page',
      description:
        'Permanently deletes a OneNote page by ID. There is no undo — confirm with the user before calling.',
      inputSchema,
      annotations: {
        destructiveHint: true,
        idempotentHint: true,
      },
    },
    async ({ pageId }) => {
      await deletePage(pageId);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ deleted: true, pageId }, null, 2),
          },
        ],
      };
    },
  );
};
