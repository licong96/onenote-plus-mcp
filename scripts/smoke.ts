/**
 * Manual end-to-end smoke test.
 *
 * Run after `onenote-mcp login` to verify every shipped capability against
 * a real OneNote account. Not wired into CI — needs real credentials.
 *
 *   ONENOTE_MCP_CLIENT_ID=<id> yarn smoke
 *
 * Idempotent. Reuses or creates a "OneNote MCP Smoke Test" notebook, and
 * cleans up every page it creates. Notebooks + sections + groups are left
 * behind for the next run.
 */

import { getAccessToken } from '@/auth/index.js';
import { createNotebook, listNotebooks } from '@/graph/notebooks.js';
import { createSection, listSections } from '@/graph/sections.js';
import {
  createSectionGroup,
  listSectionGroups,
} from '@/graph/sectionGroups.js';
import {
  createPage,
  deletePage,
  getPage,
  getPageContent,
  searchPages,
  updatePage,
} from '@/graph/pages.js';
import { htmlToMarkdown, markdownToOneNoteHtml } from '@/markdown.js';
import type { Notebook, Section, SectionGroup } from '@/graph/types.js';

const NOTEBOOK_NAME = 'OneNote MCP Smoke Test';
const SECTION_NAME = 'Smoke';
const GROUP_NAME = 'Smoke Group';
const NESTED_SECTION_NAME = 'Nested';

const TOTAL_STEPS = 13;
let stepCount = 0;
const createdPageIds: string[] = [];

const step = async <T>(label: string, fn: () => Promise<T>): Promise<T> => {
  stepCount += 1;
  const prefix = `[${stepCount}/${TOTAL_STEPS}]`;
  try {
    const result = await fn();
    process.stdout.write(`${prefix} ok   ${label}\n`);
    return result;
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    process.stderr.write(`${prefix} fail ${label}\n        → ${message}\n`);
    throw err;
  }
};

// 1×1 transparent PNG — the smallest valid bytes we can ship inline.
const TINY_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORK5CYII=',
  'base64',
);

const findOrCreateNotebook = async (name: string): Promise<Notebook> => {
  const notebooks = await listNotebooks();
  const existing = notebooks.find((n) => n.displayName === name);
  if (existing) return existing;
  return createNotebook(name);
};

const findOrCreateSection = async (
  notebookId: string,
  name: string,
): Promise<Section> => {
  const sections = await listSections(notebookId);
  const existing = sections.find((s) => s.displayName === name);
  if (existing) return existing;
  return createSection({ notebookId }, name);
};

const findOrCreateSectionGroup = async (
  notebookId: string,
  name: string,
): Promise<SectionGroup> => {
  const groups = await listSectionGroups(notebookId);
  const existing = groups.find((g) => g.displayName === name);
  if (existing) return existing;
  return createSectionGroup({ notebookId }, name);
};

const findOrCreateSectionInGroup = async (
  sectionGroupId: string,
  name: string,
): Promise<Section> => {
  // listSections() with no notebookId returns all sections flat, each carrying
  // parentSectionGroup metadata (we select+expand for it). Filter on both name
  // and parent id so we don't accumulate duplicates across runs.
  const all = await listSections();
  const existing = all.find(
    (s) => s.displayName === name && s.parentSectionGroup?.id === sectionGroupId,
  );
  if (existing) return existing;
  return createSection({ sectionGroupId }, name);
};

const assert = (cond: unknown, message: string): void => {
  if (!cond) throw new Error(`assertion failed: ${message}`);
};

const main = async (): Promise<void> => {
  await step('auth: getAccessToken()', async () => {
    const token = await getAccessToken();
    assert(token.length > 20, 'token looks too short');
  });

  await step('list_notebooks', async () => {
    const notebooks = await listNotebooks();
    assert(notebooks.length > 0, 'expected at least one notebook for the account');
  });

  const notebook = await step(
    `find or create notebook "${NOTEBOOK_NAME}"`,
    () => findOrCreateNotebook(NOTEBOOK_NAME),
  );

  const section = await step(
    `find or create section "${SECTION_NAME}"`,
    () => findOrCreateSection(notebook.id, SECTION_NAME),
  );

  const titleStamp = new Date().toISOString().replace(/[:.]/g, '-');
  const title = `smoke ${titleStamp}`;
  const created = await step(`create_page "${title}"`, async () => {
    const html = markdownToOneNoteHtml(
      '# Smoke test\n\nHello from the smoke runner.',
      title,
    );
    const page = await createPage({ sectionId: section.id, html });
    createdPageIds.push(page.id);
    return page;
  });

  await step('read_page (html → markdown round-trip)', async () => {
    const fresh = await getPage(created.id);
    assert(fresh.title === title, `expected title="${title}", got "${fresh.title}"`);
    const html = await getPageContent(created.id);
    const md = htmlToMarkdown(html);
    assert(md.includes('Hello from the smoke runner'), 'body content missing');
  });

  await step('update_page: append to body', () =>
    updatePage(created.id, [
      {
        target: 'body',
        action: 'append',
        content: '<p>appended by smoke</p>',
      },
    ]),
  );

  const newTitle = `${title} (renamed)`;
  await step('update_page: replace title (markdown → <p> unwrap)', () =>
    updatePage(created.id, [
      { target: 'title', action: 'replace', content: newTitle },
    ]),
  );

  await step(`search_pages for "${newTitle}"`, async () => {
    const hits = await searchPages({ query: newTitle, limit: 10 });
    assert(
      hits.some((p) => p.id === created.id),
      `expected to find page ${created.id} in search results`,
    );
  });

  const group = await step(
    `find or create section group "${GROUP_NAME}"`,
    () => findOrCreateSectionGroup(notebook.id, GROUP_NAME),
  );

  const nestedSection = await step(
    `create section "${NESTED_SECTION_NAME}" inside group`,
    () => findOrCreateSectionInGroup(group.id, NESTED_SECTION_NAME),
  );

  await step('create_page with PNG attachment (multipart)', async () => {
    const html = markdownToOneNoteHtml(
      `# attached\n\n![diagram](name:diagram)`,
      `${title} (attached)`,
    );
    const page = await createPage({
      sectionId: nestedSection.id,
      html,
      attachments: [
        {
          name: 'diagram',
          contentType: 'image/png',
          data: new Uint8Array(TINY_PNG),
        },
      ],
    });
    createdPageIds.push(page.id);
  });

  await step(`cleanup: delete ${createdPageIds.length} page(s)`, async () => {
    // Parallel + tolerant so one stuck delete doesn't leak the others.
    const results = await Promise.allSettled(
      createdPageIds.map((id) => deletePage(id)),
    );
    const failed = results.filter((r) => r.status === 'rejected');
    if (failed.length > 0) {
      const reasons = failed
        .map((r) => (r.status === 'rejected' ? String(r.reason) : ''))
        .filter(Boolean)
        .join('; ');
      throw new Error(
        `${failed.length}/${createdPageIds.length} delete attempts failed: ${reasons}`,
      );
    }
  });

  process.stdout.write(`\nAll ${TOTAL_STEPS} steps passed.\n`);
};

main().catch((err: unknown) => {
  if (createdPageIds.length > 0) {
    process.stderr.write(
      `\nNote: ${createdPageIds.length} test page(s) may have been left behind. IDs:\n` +
        createdPageIds.map((id) => `  ${id}`).join('\n') +
        '\n',
    );
  }
  const message = err instanceof Error ? err.stack ?? err.message : String(err);
  process.stderr.write(`\nSmoke test failed: ${message}\n`);
  process.exit(1);
});
