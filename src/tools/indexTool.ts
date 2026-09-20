import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import {
  getIndexPath,
  indexStats,
  readIndex,
  searchIndex,
  type PageIndex,
} from '@/index-store/store.js';
import { buildIndex, rebuildSkeleton, syncIndex } from '@/index-store/sync.js';
import { writeIndex } from '@/index-store/store.js';

/**
 * Index maintenance. All three modes write the same file the lookup tools read,
 * so a rebuild here is immediately visible to find_pages / resolve_page.
 */

const inputSchema = {
  mode: z
    .enum(['status', 'search', 'sync', 'rebuild', 'rebuildSkeleton'])
    .describe(
      'status: report what the index holds. search: query it offline. sync: refresh page lists (all or listed sections) keeping the skeleton. rebuild: full walk of the account. rebuildSkeleton: re-read notebooks/sections only, keeping page data for surviving sections.',
    ),
  query: z.string().optional().describe('Search text (mode: search).'),
  notebook: z.string().optional().describe('Restrict search to a notebook name (mode: search).'),
  section: z.string().optional().describe('Restrict search to a section name (mode: search).'),
  sections: z
    .array(z.string())
    .optional()
    .describe('Section IDs to refresh (mode: sync). Omit to refresh every section.'),
  limit: z.number().int().min(1).max(200).optional().describe('Max search results (default 25).'),
  concurrency: z
    .number()
    .int()
    .min(1)
    .max(10)
    .optional()
    .describe('Parallel Graph requests (default 5). Lower it if Graph throttles with 429s.'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'index',
    {
      title: 'Page Index',
      description:
        'Maintains a local mirror of the OneNote structure (notebook/section/page names and IDs) so page lookup and search are instant instead of costing a Graph round trip each. Search and status are offline. Sync/rebuild hit Graph and can take minutes on a large account — prefer sync, and scope it with `sections` when you know what changed.',
      inputSchema,
    },
    async ({ mode, query, notebook, section, sections, limit, concurrency }) => {
      const indexPath = getIndexPath();
      const existing = await readIndex(indexPath);

      const payload: Record<string, unknown> = { mode, indexPath };

      if (mode === 'status') {
        payload.indexed = Boolean(existing);
        if (existing) {
          payload.stats = indexStats(existing);
          payload.generatedAt = existing.generatedAt;
          payload.refreshedAt = existing.refreshedAt;
          payload.notebooks = existing.notebooks.map((entry) => ({
            name: entry.name,
            sectionCount: entry.sections.length,
          }));
        } else {
          payload.next = 'Run index with mode "rebuild" to create it.';
        }
        return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
      }

      if (mode === 'search') {
        if (!query || query.trim().length === 0) {
          throw new Error('mode "search" requires a non-empty query.');
        }
        const index = existing ?? (await rebuildAndPersist(indexPath, concurrency));
        const results = searchIndex(index, query, { notebook, section, limit });
        payload.query = query;
        payload.matchCount = results.length;
        payload.results = results.map((page) => ({
          id: page.id,
          title: page.title,
          notebook: page.notebookName,
          section: page.sectionName,
          groupPath: page.groupPath.length > 0 ? page.groupPath : undefined,
          lastModified: page.lastModifiedDateTime,
          webUrl: page.webUrl,
        }));
        payload.offline = true;
        return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
      }

      if (mode === 'sync') {
        if (!existing) throw new Error('No index yet. Run mode "rebuild" first.');
        const next = await syncIndex(existing, { sections, concurrency });
        await writeIndex(next, indexPath);
        payload.stats = indexStats(next);
        payload.refreshedAt = next.refreshedAt;
        payload.refreshedSections = sections?.length ?? next.notebooks.reduce((sum, n) => sum + n.sections.length, 0);
        return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
      }

      if (mode === 'rebuildSkeleton') {
        if (!existing) throw new Error('No index yet. Run mode "rebuild" first.');
        const next = await rebuildSkeleton(existing, { concurrency });
        await writeIndex(next, indexPath);
        payload.stats = indexStats(next);
        payload.generatedAt = next.generatedAt;
        payload.note = 'Skeleton refreshed; page lists were kept for sections that still exist.';
        return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
      }

      const next = await buildIndex({ concurrency });
      await writeIndex(next, indexPath);
      payload.stats = indexStats(next);
      payload.generatedAt = next.generatedAt;
      return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
    },
  );
};

const rebuildAndPersist = async (indexPath: string, concurrency?: number): Promise<PageIndex> => {
  const built = await buildIndex({ concurrency });
  await writeIndex(built, indexPath);
  return built;
};
