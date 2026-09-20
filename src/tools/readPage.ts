import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { getPage, getPageContent } from '@/graph/pages.js';
import { htmlToMarkdown } from '@/markdown.js';

const inputSchema = {
  pageId: z.string().min(1).describe('OneNote page ID to fetch.'),
  format: z
    .enum(['html', 'markdown'])
    .optional()
    .describe(
      'Content format. "html" (default) returns raw OneNote HTML; "markdown" converts it for easier consumption.',
    ),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'read_page',
    {
      title: 'Read Page',
      description:
        'Returns the full content and metadata for a single OneNote page. Defaults to HTML; pass format="markdown" for converted output.',
      inputSchema,
    },
    async ({ pageId, format }) => {
      const [meta, html] = await Promise.all([getPage(pageId), getPageContent(pageId)]);
      const content = format === 'markdown' ? htmlToMarkdown(html) : html;
      const result = {
        id: meta.id,
        title: meta.title,
        createdDateTime: meta.createdDateTime,
        lastModifiedDateTime: meta.lastModifiedDateTime,
        notebook: meta.parentNotebook?.displayName,
        section: meta.parentSection?.displayName,
        links: {
          web: meta.links?.oneNoteWebUrl?.href,
          client: meta.links?.oneNoteClientUrl?.href,
        },
        format: format ?? 'html',
        content,
      };
      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
      };
    },
  );
};
