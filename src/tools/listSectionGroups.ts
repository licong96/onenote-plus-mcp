import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { listSectionGroups } from '@/graph/sectionGroups.js';

const inputSchema = {
  notebookId: z
    .string()
    .optional()
    .describe(
      'Optional notebook ID to scope the listing. If omitted, lists section groups across all notebooks.',
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'list_section_groups',
    {
      title: 'List Section Groups',
      description:
        'Lists OneNote section groups (folders that contain sections and/or other section groups), optionally scoped to a single notebook.',
      inputSchema,
    },
    async ({ notebookId }) => {
      const groups = await listSectionGroups(notebookId);
      const summary = groups.map((g) => ({
        id: g.id,
        name: g.displayName,
        notebook: g.parentNotebook?.displayName,
        notebookId: g.parentNotebook?.id,
        parentSectionGroup: g.parentSectionGroup?.displayName,
        parentSectionGroupId: g.parentSectionGroup?.id,
        webUrl: g.links?.oneNoteWebUrl?.href,
        lastModified: g.lastModifiedDateTime,
      }));
      return {
        content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }],
      };
    },
  );
};
