// Subset of Microsoft Graph OneNote types — only the fields we surface.
// Full schema: https://learn.microsoft.com/graph/api/resources/onenote-api-overview

export interface Notebook {
  id: string;
  displayName: string;
  isDefault: boolean;
  userRole?: string;
  isShared: boolean;
  createdDateTime: string;
  lastModifiedDateTime: string;
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}

export interface SectionGroup {
  id: string;
  displayName: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  parentNotebook?: { id: string; displayName: string };
  parentSectionGroup?: { id: string; displayName: string } | null;
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}

export interface Section {
  id: string;
  displayName: string;
  isDefault: boolean;
  createdDateTime: string;
  lastModifiedDateTime: string;
  parentNotebook?: { id: string; displayName: string };
  parentSectionGroup?: { id: string; displayName: string } | null;
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}

export interface Page {
  id: string;
  title: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  contentUrl: string;
  parentSection?: { id: string; displayName: string };
  parentNotebook?: { id: string; displayName: string };
  links?: {
    oneNoteClientUrl?: { href: string };
    oneNoteWebUrl?: { href: string };
  };
}
