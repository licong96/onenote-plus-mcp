import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createSection } from '@/graph/sections.js';

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
    .describe('Parent section group ID. Provide this OR `notebookId`, exclusively.'),
  name: z
    .string()
    .min(1)
    .describe(
      "Display name for the new section. Microsoft Graph disallows the reserved characters ?*\\/:<>|&#'\"%~ and requires uniqueness within the parent.",
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'create_section',
    {
      title: 'Create Section',
      description:
        'Creates a new section inside the given notebook or section group. Exactly one of notebookId or sectionGroupId must be provided.',
      inputSchema,
    },
    async ({ notebookId, sectionGroupId, name }) => {
      if (!notebookId === !sectionGroupId) {
        throw new Error('Provide exactly one of `notebookId` or `sectionGroupId`.');
      }
      const section = await createSection({ notebookId, sectionGroupId }, name);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                id: section.id,
                name: section.displayName,
                notebook: section.parentNotebook?.displayName,
                notebookId: section.parentNotebook?.id,
                sectionGroup: section.parentSectionGroup?.displayName,
                sectionGroupId: section.parentSectionGroup?.id,
                webUrl: section.links?.oneNoteWebUrl?.href,
                createdDateTime: section.createdDateTime,
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
