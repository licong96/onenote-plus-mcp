import { graphRequest, paginate } from './client.js';
import type { Section } from './types.js';

const SECTION_SELECT =
  'id,displayName,isDefault,createdDateTime,lastModifiedDateTime,parentNotebook,parentSectionGroup,links';

const SECTION_EXPAND =
  'parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)';

export const listSections = (notebookId?: string): Promise<Section[]> => {
  const path = notebookId
    ? `/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`
    : '/me/onenote/sections';
  return paginate<Section>(path, {
    query: {
      $select: SECTION_SELECT,
      $expand: SECTION_EXPAND,
      $orderby: 'displayName',
    },
  });
};

export interface CreateSectionParent {
  notebookId?: string;
  sectionGroupId?: string;
}

export const createSection = (
  parent: CreateSectionParent,
  name: string,
): Promise<Section> => {
  const path = parent.notebookId
    ? `/me/onenote/notebooks/${encodeURIComponent(parent.notebookId)}/sections`
    : `/me/onenote/sectionGroups/${encodeURIComponent(parent.sectionGroupId!)}/sections`;
  return graphRequest<Section>(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ displayName: name }),
    // Without these, the POST response omits parent metadata and the tool's
    // summary would show undefined for the section's parent notebook / group.
    query: { $select: SECTION_SELECT, $expand: SECTION_EXPAND },
  });
};
