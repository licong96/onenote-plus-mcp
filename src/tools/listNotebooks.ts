import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { listNotebooks } from '@/graph/notebooks.js';

export const register = (server: McpServer): void => {
  server.registerTool(
    'list_notebooks',
    {
      title: 'List Notebooks',
      description:
        'Lists all OneNote notebooks accessible to the signed-in user, including notebook IDs needed for other tools.',
      inputSchema: {},
    },
    async () => {
      const notebooks = await listNotebooks();
      const summary = notebooks.map((n) => ({
        id: n.id,
        name: n.displayName,
        isDefault: n.isDefault,
        isShared: n.isShared,
        webUrl: n.links?.oneNoteWebUrl?.href,
        lastModified: n.lastModifiedDateTime,
      }));
      return {
        content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }],
      };
    },
  );
};
