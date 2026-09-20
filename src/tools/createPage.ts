import { readFile, realpath } from 'node:fs/promises';
import { isAbsolute, relative, resolve, sep } from 'node:path';
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createPage, type CreatePageAttachment } from '@/graph/pages.js';
import { htmlToOneNotePage, markdownToOneNoteHtml } from '@/markdown.js';

const PART_NAME_REGEX = /^[A-Za-z0-9._-]+$/;
const BASE64_REGEX = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;

const attachmentSchema = z
  .object({
    name: z
      .string()
      .min(1)
      .regex(PART_NAME_REGEX, 'must be alphanumeric, dots, underscores, or hyphens')
      .describe(
        'Unique part name (alphanumeric/dot/underscore/hyphen). Reference it from the content as `name:<name>` (e.g. `<img src="name:diagram1" />`).',
      ),
    contentType: z.string().min(1).describe('MIME type, e.g. `image/png`, `application/pdf`.'),
    path: z
      .string()
      .min(1)
      .optional()
      .describe(
        "Local file path, relative or absolute. Must resolve inside the server's working directory (path traversal is rejected).",
      ),
    data: z
      .string()
      .min(1)
      .regex(BASE64_REGEX, 'must be valid base64')
      .optional()
      .describe('Base64-encoded bytes. Use this OR `path`.'),
  })
  .refine((a) => Boolean(a.path) !== Boolean(a.data), {
    message: 'Provide exactly one of `path` or `data`.',
    path: ['path'],
  });

const escapesCwd = (cwd: string, target: string): boolean => {
  const rel = relative(cwd, target);
  return rel === '..' || rel.startsWith(`..${sep}`) || isAbsolute(rel);
};

export const readLocalAttachment = async (rawPath: string): Promise<Uint8Array> => {
  const cwd = await realpath(process.cwd());
  const full = resolve(cwd, rawPath);
  // Reject paths that escape cwd via "..", an absolute path elsewhere, or symlink
  // targets. This is critical when the tool is invoked by an LLM that could be
  // coaxed into exfiltrating local secrets (~/.ssh/id_rsa, /etc/passwd, …) by
  // attaching them to a OneNote page.
  const escaped = (): never => {
    throw new Error(
      `Attachment path "${rawPath}" resolves outside the working directory; refusing to read.`,
    );
  };
  if (escapesCwd(cwd, full)) escaped();
  // resolve() is purely lexical, so re-check after following symlinks.
  const real = await realpath(full);
  if (escapesCwd(cwd, real)) escaped();
  return new Uint8Array(await readFile(real));
};

const inputSchema = {
  sectionId: z.string().min(1).describe('Section ID to create the page in.'),
  title: z.string().min(1).describe('Page title (used in the <title> element).'),
  content: z.string().min(1).describe('Page body. Format is determined by the `format` field.'),
  format: z
    .enum(['markdown', 'html'])
    .default('markdown')
    .describe(
      'Content format. "markdown" (default) is converted to HTML; "html" is sent directly (wrapped if it is a body fragment).',
    ),
  attachments: z
    .array(attachmentSchema)
    .optional()
    .describe(
      'Optional binary attachments. Reference each from `content` via `name:<name>` URIs (e.g. `![alt](name:diagram1)` in markdown, or `<img src="name:diagram1" />` / `<object data="name:file1" data-attachment="readme.pdf" type="application/pdf" />` in HTML).',
    ),
};

const loadAttachments = (
  parts: z.infer<typeof attachmentSchema>[],
): Promise<CreatePageAttachment[]> =>
  Promise.all(
    parts.map(async (a) => {
      const data = a.path
        ? await readLocalAttachment(a.path)
        : new Uint8Array(Buffer.from(a.data ?? '', 'base64'));
      return { name: a.name, contentType: a.contentType, data };
    }),
  );

export const register = (server: McpServer): void => {
  server.registerTool(
    'create_page',
    {
      title: 'Create Page',
      description:
        'Creates a new OneNote page in the given section. Accepts Markdown (default) or HTML; optional `attachments` upload binary parts referenced from the content via `name:<name>` URIs.',
      inputSchema,
    },
    async ({ sectionId, title, content, format, attachments }) => {
      const html =
        format === 'html'
          ? htmlToOneNotePage(content, title)
          : markdownToOneNoteHtml(content, title);
      const parts = attachments && attachments.length > 0
        ? await loadAttachments(attachments)
        : undefined;
      const page = await createPage({ sectionId, html, attachments: parts });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                id: page.id,
                title: page.title,
                createdDateTime: page.createdDateTime,
                webUrl: page.links?.oneNoteWebUrl?.href,
                attachmentCount: parts?.length ?? 0,
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
