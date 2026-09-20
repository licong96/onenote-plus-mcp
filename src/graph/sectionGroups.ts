import { graphRequest, paginate } from './client.js';
import type { SectionGroup } from './types.js';

const SECTION_GROUP_SELECT =
  'id,displayName,createdDateTime,lastModifiedDateTime,parentNotebook,parentSectionGroup,links';

const SECTION_GROUP_EXPAND =
  'parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)';

export const listSectionGroups = (notebookId?: string): Promise<SectionGroup[]> => {
  const path = notebookId
    ? `/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sectionGroups`
    : '/me/onenote/sectionGroups';
  return paginate<SectionGroup>(path, {
    query: {
      $select: SECTION_GROUP_SELECT,
      $expand: SECTION_GROUP_EXPAND,
      $orderby: 'displayName',
    },
  });
};

export interface CreateSectionGroupParent {
  notebookId?: string;
  sectionGroupId?: string;
}

export const createSectionGroup = (
  parent: CreateSectionGroupParent,
  name: string,
): Promise<SectionGroup> => {
  const path = parent.notebookId
    ? `/me/onenote/notebooks/${encodeURIComponent(parent.notebookId)}/sectionGroups`
    : `/me/onenote/sectionGroups/${encodeURIComponent(parent.sectionGroupId!)}/sectionGroups`;
  return graphRequest<SectionGroup>(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ displayName: name }),
    // Without $expand the POST response would omit parent metadata and the
    // tool's summary would show undefined for the parent notebook / group name.
    query: { $select: SECTION_GROUP_SELECT, $expand: SECTION_GROUP_EXPAND },
  });
};
