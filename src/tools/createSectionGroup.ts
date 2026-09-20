import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createSectionGroup } from '@/graph/sectionGroups.js';

const inputSchema = {
  notebookId: z
    .string()
    .min(1)
    .optional()
    .describe('Parent notebook ID. Provide this OR `sectionGroupId`, exclusively.'),
  sectionGroupId: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Parent section group ID (for nested section groups). Provide this OR `notebookId`, exclusively.',
    ),
  name: z
    .string()
    .min(1)
    .describe(
      "Display name for the new section group. Same character restrictions as sections (no ?*\\/:<>|&#'\"%~).",
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'create_section_group',
    {
      title: 'Create Section Group',
      description:
        'Creates a new section group inside the given notebook or another section group. Exactly one of notebookId or sectionGroupId must be provided.',
      inputSchema,
    },
    async ({ notebookId, sectionGroupId, name }) => {
      if (!notebookId === !sectionGroupId) {
        throw new Error('Provide exactly one of `notebookId` or `sectionGroupId`.');
      }
      const group = await createSectionGroup({ notebookId, sectionGroupId }, name);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                id: group.id,
                name: group.displayName,
                notebook: group.parentNotebook?.displayName,
                notebookId: group.parentNotebook?.id,
                parentSectionGroup: group.parentSectionGroup?.displayName,
                parentSectionGroupId: group.parentSectionGroup?.id,
                webUrl: group.links?.oneNoteWebUrl?.href,
                createdDateTime: group.createdDateTime,
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
