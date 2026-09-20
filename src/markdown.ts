import { marked } from 'marked';
import TurndownService from 'turndown';

const turndown = new TurndownService({
  headingStyle: 'atx',
  bulletListMarker: '-',
  codeBlockStyle: 'fenced',
  emDelimiter: '_',
});

const escapeHtml = (input: string): string =>
  input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

/**
 * Convert Markdown into a full HTML document suitable for the OneNote
 * `POST /sections/{id}/pages` endpoint, which expects an XHTML page with
 * <html>, <head><title>…</title></head>, and <body>.
 */
export const markdownToOneNoteHtml = (markdown: string, title: string): string => {
  const bodyHtml = marked.parse(markdown, { async: false }) as string;
  const safeTitle = escapeHtml(title);
  return `<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>${safeTitle}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
${bodyHtml}
  </body>
</html>`;
};

/**
 * Wrap raw HTML body content in a OneNote-compatible page document.
 * Idempotent: if `html` already contains <html>…</html>, returns it unchanged.
 */
export const htmlToOneNotePage = (html: string, title: string): string => {
  if (/<html[\s>]/i.test(html)) return html;
  const safeTitle = escapeHtml(title);
  return `<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <title>${safeTitle}</title>
    <meta name="created" content="${new Date().toISOString()}" />
  </head>
  <body>
${html}
  </body>
</html>`;
};

export const htmlToMarkdown = (html: string): string => turndown.turndown(html);

/**
 * Convert Markdown into an HTML fragment (no <html>/<body> wrapper) for use
 * as the `content` of an `update_page` PATCH command.
 */
export const markdownToHtmlFragment = (markdown: string): string =>
  (marked.parse(markdown, { async: false }) as string).trim();
