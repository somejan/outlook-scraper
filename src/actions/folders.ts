/**
 * Folder enumeration via OWA `FindFolder` / `GetFolder`.
 *
 * The request shape mirrors what the new Outlook web app actually sends (see
 * README "Verifying the API"): a `FolderResponseShape` with `BaseShape: IdOnly`
 * plus explicit `AdditionalProperties`, and an `IndexedPageView` — sending
 * `BaseShape: Default` without paging triggers an OwaSerializationException.
 */

import type { OwaClient } from "../owa-client.js";
import type { DistinguishedFolder, MailFolder } from "../types.js";
import { folderIdFragment } from "./ews.js";

/** Folder fields we ask the server to return. */
const FOLDER_PROPERTIES = [
  "DistinguishedFolderId",
  "FolderClass",
  "FolderDisplayName",
  "ParentFolderId",
  "TotalCount",
  "UnreadCount",
].map((FieldURI) => ({ __type: "PropertyUri:#Exchange", FieldURI }));

function folderShape() {
  return {
    __type: "FolderResponseShape:#Exchange",
    BaseShape: "IdOnly",
    AdditionalProperties: FOLDER_PROPERTIES,
  };
}

/**
 * List the child folders under a parent (default: the mailbox root, which
 * yields Inbox, Sent Items, etc.). Uses a `Deep` traversal so nested folders
 * are included.
 */
export async function findFolders(
  client: OwaClient,
  parent: string = "msgfolderroot"
): Promise<MailFolder[]> {
  const json = await client.call("FindFolder", {
    Traversal: "Deep",
    FolderShape: folderShape(),
    Paging: {
      __type: "IndexedPageView:#Exchange",
      BasePoint: "Beginning",
      Offset: 0,
      MaxEntriesReturned: 1000,
    },
    ParentFolderIds: [folderIdFragment(parent)],
  });

  const rm = json?.Body?.ResponseMessages?.Items?.[0];
  const folders: any[] = rm?.RootFolder?.Folders ?? [];
  return folders.map(parseFolder);
}

/** Get a single folder's metadata by distinguished name or id. */
export async function getFolder(
  client: OwaClient,
  folder: string
): Promise<MailFolder> {
  const json = await client.call("GetFolder", {
    FolderShape: folderShape(),
    FolderIds: [folderIdFragment(folder)],
  });
  const rm = json?.Body?.ResponseMessages?.Items?.[0];
  const raw = rm?.Folders?.[0];
  return parseFolder(raw);
}

function parseFolder(raw: any): MailFolder {
  // DistinguishedFolderId may come back as a string or as an object { Id }.
  const dfid = raw?.DistinguishedFolderId;
  const distinguished = (typeof dfid === "string" ? dfid : dfid?.Id) as
    | DistinguishedFolder
    | undefined;
  return {
    id: raw?.FolderId?.Id ?? "",
    displayName:
      raw?.DisplayName ?? raw?.FolderDisplayName ?? distinguished ?? "(unnamed)",
    distinguished,
    totalCount: typeof raw?.TotalCount === "number" ? raw.TotalCount : undefined,
    unreadCount:
      typeof raw?.UnreadCount === "number" ? raw.UnreadCount : undefined,
    parentId: raw?.ParentFolderId?.Id,
  };
}
