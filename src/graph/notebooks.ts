import { graphRequest, paginate } from './client.js';
import type { Notebook } from './types.js';

const NOTEBOOK_SELECT =
  'id,displayName,isDefault,userRole,isShared,createdDateTime,lastModifiedDateTime,links';

export const listNotebooks = (): Promise<Notebook[]> =>
  paginate<Notebook>('/me/onenote/notebooks', {
    query: {
      $select: NOTEBOOK_SELECT,
      $orderby: 'displayName',
    },
  });

export const createNotebook = (name: string): Promise<Notebook> =>
  graphRequest<Notebook>('/me/onenote/notebooks', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ displayName: name }),
    // $select on POST shapes the response so callers don't see undefined fields.
    query: { $select: NOTEBOOK_SELECT },
  });
