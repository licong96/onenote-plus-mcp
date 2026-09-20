import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { listSections } from '@/graph/sections.js';

const inputSchema = {
  notebookId: z
    .string()
    .optional()
    .describe(
      'Optional notebook ID to scope the listing. If omitted, lists sections across all notebooks.',
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'list_sections',
    {
      title: 'List Sections',
      description:
        'Lists OneNote sections, optionally scoped to a single notebook by ID. Section IDs are needed by create_page.',
      inputSchema,
    },
    async ({ notebookId }) => {
      const sections = await listSections(notebookId);
      const summary = sections.map((s) => ({
        id: s.id,
        name: s.displayName,
        isDefault: s.isDefault,
        notebook: s.parentNotebook?.displayName,
        notebookId: s.parentNotebook?.id,
        webUrl: s.links?.oneNoteWebUrl?.href,
        lastModified: s.lastModifiedDateTime,
      }));
      return {
        content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }],
      };
    },
  );
};
