import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { updatePage, type UpdatePageCommand } from '@/graph/pages.js';
import { markdownToHtmlFragment } from '@/markdown.js';

const operationSchema = z
  .object({
    target: z
      .string()
      .min(1)
      .describe(
        'The element to act on: `body`, `title`, or the raw `data-id` of an element from read_page HTML. A leading `#` is accepted and stripped (so CSS-style selectors copied from elsewhere also work).',
      ),
    action: z
      .enum(['append', 'prepend', 'insert', 'replace', 'delete'])
      .describe(
        'append/prepend add content inside the target; insert places content next to the target (set position); replace swaps the target; delete removes it.',
      ),
    position: z
      .enum(['before', 'after'])
      .optional()
      .describe('Required when `action` is `insert`. Determines which side of the target.'),
    content: z
      .string()
      .optional()
      .describe('Content body. Required for every action except `delete`.'),
    format: z
      .enum(['markdown', 'html'])
      .default('markdown')
      .describe('Format of `content`. Markdown (default) is converted to HTML before sending.'),
  })
  .refine((op) => op.action === 'delete' || op.content !== undefined, {
    message: '`content` is required for every action except `delete`.',
    path: ['content'],
  })
  .refine((op) => op.action !== 'insert' || op.position !== undefined, {
    message: '`position` is required when `action` is `insert`.',
    path: ['position'],
  });

const inputSchema = {
  pageId: z.string().min(1).describe('OneNote page ID to update.'),
  operations: z
    .array(operationSchema)
    .min(1)
    .describe('Ordered list of edits to apply to the page in a single request.'),
};

const isTitleTarget = (target: string): boolean =>
  (target.startsWith('#') ? target.slice(1) : target) === 'title';

const toGraphCommand = (op: z.infer<typeof operationSchema>): UpdatePageCommand => {
  const cmd: UpdatePageCommand = {
    target: op.target,
    action: op.action,
  };
  if (op.position) cmd.position = op.position;
  if (op.content !== undefined) {
    let content = op.format === 'html' ? op.content : markdownToHtmlFragment(op.content);
    // marked wraps single-line content in <p>…</p>. The title element renders
    // those literally, so unwrap exactly one such pair when targeting the title.
    if (isTitleTarget(op.target) && op.format !== 'html') {
      const unwrapped = content.match(/^<p>([\s\S]*)<\/p>$/);
      if (unwrapped) content = unwrapped[1]!;
    }
    cmd.content = content;
  }
  return cmd;
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'update_page',
    {
      title: 'Update Page',
      description:
        'Applies one or more edits to an existing OneNote page (append, prepend, insert, replace, delete). Targets are `body`, `title`, or `#<data-id>` selectors — find data-id values in the HTML returned by read_page.',
      inputSchema,
      annotations: {
        destructiveHint: true,
        idempotentHint: false,
      },
    },
    async ({ pageId, operations }) => {
      const commands = operations.map(toGraphCommand);
      await updatePage(pageId, commands);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              { updated: true, pageId, operationCount: commands.length },
              null,
              2,
            ),
          },
        ],
      };
    },
  );
};
