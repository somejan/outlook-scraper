/**
 * Message listing and reading via OWA `FindItem` / `GetItem`.
 */

import type { OwaClient } from "../owa-client.js";
import type { ListOptions, MailMessage, MailSummary } from "../types.js";
import {
  findItemResults,
  folderIdFragment,
  getItemResult,
  itemIdFragment,
  parseAddressList,
  parseAttachments,
  parseSummary,
} from "./ews.js";

/**
 * List messages in a folder, newest first. Returns lightweight summaries
 * (no body); call {@link getMessage} for the full message.
 */
export async function findItems(
  client: OwaClient,
  folder: string,
  options: ListOptions = {}
): Promise<MailSummary[]> {
  const limit = options.limit ?? 50;
  const offset = options.offset ?? 0;

  const json = await client.call("FindItem", {
    ItemShape: {
      __type: "ItemResponseShape:#Exchange",
      BaseShape: "IdOnly",
      // Request the cheap summary fields explicitly so FindItem stays fast.
      AdditionalProperties: summaryProperties(),
    },
    ParentFolderIds: [folderIdFragment(folder)],
    Traversal: "Shallow",
    Paging: {
      __type: "IndexedPageView:#Exchange",
      BasePoint: "Beginning",
      MaxEntriesReturned: limit,
      Offset: offset,
    },
    SortOrder: [
      {
        __type: "SortResults:#Exchange",
        Order: "Descending",
        Path: {
          __type: "PropertyUri:#Exchange",
          FieldURI: "item:DateTimeReceived",
        },
      },
    ],
  });

  return findItemResults(json).map(parseSummary);
}

/**
 * Load a full message by id: recipients, body, and attachment descriptors.
 * When `includeMime` is set, also request the raw RFC-822 MIME (used by export).
 */
export async function getItem(
  client: OwaClient,
  id: string,
  changeKey?: string,
  includeMime = false
): Promise<MailMessage> {
  const json = await client.call("GetItem", {
    ItemShape: {
      __type: "ItemResponseShape:#Exchange",
      BaseShape: "Default",
      BodyType: "HTML",
      IncludeMimeContent: includeMime,
      AdditionalProperties: summaryProperties(),
    },
    ItemIds: [itemIdFragment(id, changeKey)],
  });

  const item = getItemResult(json);
  const summary = parseSummary(item);
  const bodyNode = item?.Body;
  const message: MailMessage = {
    ...summary,
    to: parseAddressList(item?.ToRecipients),
    cc: parseAddressList(item?.CcRecipients),
    bcc: parseAddressList(item?.BccRecipients),
    body: bodyNode
      ? {
          contentType: bodyNode.BodyType === "Text" ? "Text" : "HTML",
          content: bodyNode.Value ?? "",
        }
      : undefined,
    attachments: parseAttachments(item),
  };

  if (includeMime && item?.MimeContent?.Value) {
    // MimeContent.Value is base64-encoded RFC-822.
    message.mime = Buffer.from(item.MimeContent.Value, "base64").toString("utf8");
  }
  return message;
}

/** Summary properties requested on both FindItem and GetItem. */
function summaryProperties(): unknown[] {
  const fields = [
    "item:Subject",
    "item:DateTimeReceived",
    "item:DateTimeSent",
    "item:Size",
    "item:HasAttachments",
    "message:From",
    "message:Sender",
    "message:IsRead",
  ];
  return fields.map((FieldURI) => ({
    __type: "PropertyUri:#Exchange",
    FieldURI,
  }));
}
