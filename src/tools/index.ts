import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { register as registerListNotebooks } from './listNotebooks.js';
import { register as registerListSections } from './listSections.js';
import { register as registerSearchPages } from './searchPages.js';
import { register as registerReadPage } from './readPage.js';
import { register as registerCreatePage } from './createPage.js';
import { register as registerUpdatePage } from './updatePage.js';
import { register as registerDeletePage } from './deletePage.js';
import { register as registerCreateNotebook } from './createNotebook.js';
import { register as registerCreateSection } from './createSection.js';
import { register as registerListSectionGroups } from './listSectionGroups.js';
import { register as registerCreateSectionGroup } from './createSectionGroup.js';

export const registerAllTools = (server: McpServer): void => {
  registerListNotebooks(server);
  registerListSections(server);
  registerListSectionGroups(server);
  registerSearchPages(server);
  registerReadPage(server);
  registerCreateNotebook(server);
  registerCreateSection(server);
  registerCreateSectionGroup(server);
  registerCreatePage(server);
  registerUpdatePage(server);
  registerDeletePage(server);
};
