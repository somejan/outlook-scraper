/**
 * Shared helpers for building OWA/EWS request fragments and parsing the common
 * shapes that appear across `FindItem`, `GetItem`, `FindFolder`, etc.
 *
 * Keeping the id/address/date plumbing here means each action file stays focused
 * on its own request/response and the fiddly parsing is unit-testable in one
 * place (see `test/restriction.test.ts`).
 */

import type {
  Attachment,
  DistinguishedFolder,
  EmailAddress,
  MailSummary,
} from "../types.js";

/** The set of distinguished folder names OWA accepts as a parent id. */
const DISTINGUISHED: readonly DistinguishedFolder[] = [
  "inbox",
  "sentitems",
  "drafts",
  "deleteditems",
  "junkemail",
  "archive",
  "outbox",
  "msgfolderroot",
];

/** Whether a string is a well-known distinguished folder name. */
export function isDistinguished(name: string): name is DistinguishedFolder {
  return (DISTINGUISHED as readonly string[]).includes(name.toLowerCase());
}

/**
 * Build a `ParentFolderId`/`FolderId` fragment. A distinguished name (e.g.
 * `inbox`) becomes a `DistinguishedFolderId`; anything else is treated as an
 * opaque `FolderId`.
 */
export function folderIdFragment(folder: string): Record<string, unknown> {
  if (isDistinguished(folder)) {
    return {
      __type: "DistinguishedFolderId:#Exchange",
      Id: folder.toLowerCase(),
    };
  }
  return { __type: "FolderId:#Exchange", Id: folder };
}

/** Build an `ItemId` fragment from an opaque item id (+ optional change key). */
export function itemIdFragment(id: string, changeKey?: string): Record<string, unknown> {
  const frag: Record<string, unknown> = { __type: "ItemId:#Exchange", Id: id };
  if (changeKey) frag.ChangeKey = changeKey;
  return frag;
}

/** Parse a single EWS mailbox/recipient node into an {@link EmailAddress}. */
export function parseAddress(node: any): EmailAddress | undefined {
  const mbx = node?.Mailbox ?? node;
  if (!mbx) return undefined;
  const out: EmailAddress = {};
  if (mbx.Name) out.name = mbx.Name;
  if (mbx.EmailAddress) out.address = mbx.EmailAddress;
  return out.name || out.address ? out : undefined;
}

/** Parse an array of recipient nodes (To/Cc/Bcc). */
export function parseAddressList(list: any): EmailAddress[] {
  if (!Array.isArray(list)) return [];
  return list.map((n) => parseAddress(n)).filter((a): a is EmailAddress => !!a);
}

/** Parse an EWS date-time string into a `Date`, tolerating missing values. */
export function parseDate(value: unknown): Date | undefined {
  if (typeof value !== "string" || !value) return undefined;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

/** Parse the shared "message summary" fields present on a `Message` item. */
export function parseSummary(item: any): MailSummary {
  return {
    id: item?.ItemId?.Id ?? "",
    changeKey: item?.ItemId?.ChangeKey,
    subject: item?.Subject ?? "(no subject)",
    from: parseAddress(item?.From) ?? parseAddress(item?.Sender),
    preview: item?.Preview ?? undefined,
    received: parseDate(item?.DateTimeReceived ?? item?.DateTimeSent),
    isRead: typeof item?.IsRead === "boolean" ? item.IsRead : undefined,
    hasAttachments:
      typeof item?.HasAttachments === "boolean" ? item.HasAttachments : undefined,
    size: typeof item?.Size === "number" ? item.Size : undefined,
  };
}

/** Parse the attachment descriptors attached to a `GetItem` message. */
export function parseAttachments(item: any): Attachment[] {
  const raw: any[] = item?.Attachments ?? [];
  return raw
    .map((a): Attachment | undefined => {
      const id = a?.AttachmentId?.Id;
      if (!id) return undefined;
      return {
        id,
        name: a.Name ?? "attachment",
        contentType: a.ContentType,
        size: typeof a.Size === "number" ? a.Size : undefined,
        isInline: a.IsInline === true,
      };
    })
    .filter((a): a is Attachment => !!a);
}

/** Pull the list of items out of a `FindItem` response envelope. */
export function findItemResults(json: any): any[] {
  const rm = json?.Body?.ResponseMessages?.Items?.[0];
  return rm?.RootFolder?.Items ?? [];
}

/** Pull the single item out of a `GetItem` response envelope. */
export function getItemResult(json: any): any {
  const rm = json?.Body?.ResponseMessages?.Items?.[0];
  return rm?.Items?.[0];
}
