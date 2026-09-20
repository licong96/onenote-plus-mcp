import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { copyPageToSection } from '@/graph/pages.js';
import { getPage } from '@/graph/pages.js';
import { getSection } from '@/graph/sections.js';
import { afterPageCopied } from '@/index-store/maintain.js';

const inputSchema = {
  pageId: z.string().min(1).describe('ID of the page to copy.'),
  targetSectionId: z.string().min(1).describe('ID of the section to copy the page into.'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'copy_page',
    {
      title: 'Copy Page',
      description:
        'Copies a page into another section, alongside its original. The source page is left untouched — this never moves or deletes anything.',
      inputSchema,
    },
    async ({ pageId, targetSectionId }) => {
      // Resolve both ends first so the response names real things instead of
      // raw GUIDs, and so a bad ID fails before an async copy is queued.
      const [page, target] = await Promise.all([getPage(pageId), getSection(targetSectionId)]);
      const result = await copyPageToSection(pageId, targetSectionId);

      // The copy is async (202), so the target section may not list the new page
      // yet. Refresh anyway: it is idempotent, and a later lookup that misses
      // will repair itself via resolve_page.
      const indexUpdate = await afterPageCopied(targetSectionId);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                copied: true,
                status: result.status,
                operationUrl: result.operationUrl,
                page: { id: page.id, title: page.title },
                into: {
                  id: target.id,
                  name: target.displayName,
                  notebook: target.parentNotebook?.displayName,
                },
                indexUpdate,
                note:
                  result.status === 202
                    ? 'Graph accepted the copy (202) and finishes it asynchronously — the new page appears in the target section shortly.'
                    : 'Copy request completed synchronously.',
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
