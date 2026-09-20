import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createNotebook } from '@/graph/notebooks.js';

const inputSchema = {
  name: z
    .string()
    .min(1)
    .describe(
      "Display name for the new notebook. Microsoft Graph requires this to be unique within the user's notebooks and to avoid the reserved characters ?*\\/:<>|&#'\"%~.",
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'create_notebook',
    {
      title: 'Create Notebook',
      description: 'Creates a new top-level OneNote notebook with the given display name.',
      inputSchema,
    },
    async ({ name }) => {
      const notebook = await createNotebook(name);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                id: notebook.id,
                name: notebook.displayName,
                isDefault: notebook.isDefault,
                webUrl: notebook.links?.oneNoteWebUrl?.href,
                createdDateTime: notebook.createdDateTime,
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
