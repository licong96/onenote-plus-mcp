import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { findPages } from '@/graph/find.js';

const inputSchema = {
  query: z
    .string()
    .min(1)
    .describe('Text to look for. Case-insensitive substring match, so plain words work — no operators, wildcards, or quotes.'),
  sectionId: z.string().optional().describe('Limit the search to one section (fastest).'),
  notebookId: z.string().optional().describe('Limit the search to one notebook.'),
  limit: z.number().int().min(1).max(200).optional().describe('Maximum matches to return (default 25).'),
  includeContent: z
    .boolean()
    .optional()
    .describe(
      'Also search page bodies, not just titles. Much slower: it downloads each page. Off by default.',
    ),
  maxContentPages: z
    .number()
    .int()
    .min(1)
    .max(300)
    .optional()
    .describe('Cap on pages downloaded for the content pass (default 100, max 300).'),
};

export const register = (server: McpServer): void => {
  server.registerTool(
    'find_pages',
    {
      title: 'Find Pages',
      description:
        'Searches OneNote pages by text and returns matches with page IDs, titles, and where they live. Works section-scoped, notebook-scoped, or across the whole account. This is the working alternative to search_pages, which relies on a Graph search endpoint that rejects requests on this account.',
      inputSchema,
    },
    async ({ query, sectionId, notebookId, limit, includeContent, maxContentPages }) => {
      const result = await findPages({
        query,
        sectionId,
        notebookId,
        limit,
        includeContent,
        maxContentPages,
      });
      return {
        content: [{ type: 'text', text: JSON.stringify({ query, ...result }, null, 2) }],
      };
    },
  );
};
