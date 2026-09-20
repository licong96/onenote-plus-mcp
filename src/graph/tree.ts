import { paginate } from './client.js';
import { DEFAULT_CONCURRENCY, mapWithConcurrency } from '@/util/concurrency.js';
import type { Notebook, Section, SectionGroup } from './types.js';

// No $orderby on these collections on purpose: the OneNote API accepts it on
// some endpoints and rejects the whole request with "unsupported OData query
// parameters" on others. These lists are small, so we sort client-side instead.
const SECTION_SELECT = 'id,displayName,isDefault,createdDateTime,lastModifiedDateTime';
const GROUP_SELECT = 'id,displayName,createdDateTime,lastModifiedDateTime';
const PAGES_COUNT_SELECT = 'id';

/** Guard against pathological nesting; OneNote itself allows only a few levels. */
const MAX_GROUP_DEPTH = 5;

export interface SectionNode {
  id: string;
  name: string;
  isDefault: boolean;
  lastModified: string;
  /** Section-group names above this section, outermost first. Empty for top level. */
  groupPath: string[];
  /** Page count, only populated when the caller asks for it. */
  pageCount?: number;
}

export interface NotebookNode {
  id: string;
  name: string;
  isDefault: boolean;
  isShared: boolean;
  lastModified: string;
  webUrl?: string;
  sections: SectionNode[];
}

const toNode = (section: Section, groupPath: string[]): SectionNode => ({
  id: section.id,
  name: section.displayName,
  isDefault: section.isDefault,
  lastModified: section.lastModifiedDateTime,
  groupPath,
});

export const listTopLevelSections = (notebookId: string): Promise<Section[]> =>
  paginate<Section>(`/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`, {
    query: { $select: SECTION_SELECT },
  });

export const listTopLevelSectionGroups = (notebookId: string): Promise<SectionGroup[]> =>
  paginate<SectionGroup>(`/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sectionGroups`, {
    query: { $select: GROUP_SELECT },
  });

export const listGroupSections = (groupId: string): Promise<Section[]> =>
  paginate<Section>(`/me/onenote/sectionGroups/${encodeURIComponent(groupId)}/sections`, {
    query: { $select: SECTION_SELECT },
  });

export const listChildSectionGroups = (groupId: string): Promise<SectionGroup[]> =>
  paginate<SectionGroup>(`/me/onenote/sectionGroups/${encodeURIComponent(groupId)}/sectionGroups`, {
    query: { $select: GROUP_SELECT },
  });

const collectGroupSections = async (
  group: SectionGroup,
  parentPath: string[],
  depth: number,
  out: SectionNode[],
): Promise<void> => {
  if (depth > MAX_GROUP_DEPTH) return;
  const path = [...parentPath, group.displayName];
  const [sections, children] = await Promise.all([
    listGroupSections(group.id),
    listChildSectionGroups(group.id),
  ]);
  for (const section of sections) out.push(toNode(section, path));
  for (const child of children) {
    await collectGroupSections(child, path, depth + 1, out);
  }
};

/**
 * Every section of a notebook, flattened, with the section-group path each one
 * lives under. `/notebooks/{id}/sections` alone misses sections nested in
 * section groups, which silently truncates any notebook-wide scan built on it.
 */
export const collectSectionsForNotebook = async (notebookId: string): Promise<SectionNode[]> => {
  const [topLevel, groups] = await Promise.all([
    listTopLevelSections(notebookId),
    listTopLevelSectionGroups(notebookId),
  ]);

  const nodes: SectionNode[] = topLevel.map((section) => toNode(section, []));
  for (const group of groups) {
    await collectGroupSections(group, [], 0, nodes);
  }

  return nodes.sort((a, b) => a.name.localeCompare(b.name));
};

export const listNotebooksRaw = (): Promise<Notebook[]> =>
  paginate<Notebook>('/me/onenote/notebooks', {
    query: {
      $select: 'id,displayName,isDefault,userRole,isShared,createdDateTime,lastModifiedDateTime,links',
    },
  });

export const getNotebook = (notebookId: string): Promise<Notebook> =>
  paginate<Notebook>('/me/onenote/notebooks', {
    query: {
      $select: 'id,displayName,isDefault,userRole,isShared,createdDateTime,lastModifiedDateTime,links',
      $filter: `id eq '${notebookId.replaceAll("'", "''")}'`,
    },
  }).then((rows) => {
    const found = rows[0];
    if (!found) throw new Error(`Notebook not found: ${notebookId}`);
    return found;
  });

/**
 * Full hierarchy for one notebook, or for every notebook when `notebookId` is
 * omitted. Notebooks are scanned concurrently — a 13-notebook account is ~90
 * requests, which is slow sequentially and throttled when fully parallel.
 */
export const collectNotebookTree = async (
  notebookId?: string,
  concurrency: number = DEFAULT_CONCURRENCY,
): Promise<NotebookNode[]> => {
  const notebooks = notebookId ? [await getNotebook(notebookId)] : await listNotebooksRaw();

  return mapWithConcurrency(notebooks, concurrency, async (notebook) => ({
    id: notebook.id,
    name: notebook.displayName,
    isDefault: notebook.isDefault,
    isShared: notebook.isShared,
    lastModified: notebook.lastModifiedDateTime,
    webUrl: notebook.links?.oneNoteWebUrl?.href,
    sections: await collectSectionsForNotebook(notebook.id),
  }));
};

/**
 * Count pages in a section.
 *
 * The API exposes no `$count` here (it answers "unsupported OData query
 * parameters" on `/me/onenote/pages`), so this walks the collection one page
 * per row. Opt-in only: it costs one request per 100 pages, per section.
 */
export const countPagesInSection = async (sectionId: string): Promise<number> => {
  const pages = await paginate<{ id: string }>(
    `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages`,
    { query: { $select: PAGES_COUNT_SELECT, $top: 100 } },
  );
  return pages.length;
};

export const countPagesForSections = async (
  sections: readonly SectionNode[],
  concurrency: number = DEFAULT_CONCURRENCY,
): Promise<SectionNode[]> =>
  mapWithConcurrency(sections, concurrency, async (section) => ({
    ...section,
    pageCount: await countPagesInSection(section.id),
  }));
