/**
 * Shared type definitions for the Outlook (OWA) scraper library.
 */

/** Strategy used to talk to Outlook Web. `"rest"` is the only implemented path today. */
export type Strategy = "auto" | "rest" | "dom";

/** Well-known "distinguished" mail folders addressable by name. */
export type DistinguishedFolder =
  | "inbox"
  | "sentitems"
  | "drafts"
  | "deleteditems"
  | "junkemail"
  | "archive"
  | "outbox"
  | "msgfolderroot";

/**
 * An EWS ItemId / FolderId. OWA identifies items and folders with an opaque,
 * change-tracking id string. `changeKey` is optional and only needed by a few
 * mutating actions (we are read-only, so it is informational).
 */
export interface EwsId {
  id: string;
  changeKey?: string;
}

/** A mail folder in the mailbox. */
export interface MailFolder {
  /** Opaque folder id used by the OWA API. */
  id: string;
  /** Display name, e.g. `Inbox` / `Bandeja de entrada`. */
  displayName: string;
  /** Distinguished-folder key when this is a well-known folder. */
  distinguished?: DistinguishedFolder;
  /** Total number of items in the folder, when reported. */
  totalCount?: number;
  /** Number of unread items, when reported. */
  unreadCount?: number;
  /** Parent folder id, when reported. */
  parentId?: string;
}

/** A single attachment on a message. */
export interface Attachment {
  /** Opaque attachment id used by the OWA API (`GetAttachment`). */
  id: string;
  /** File name as shown in the message. */
  name: string;
  /** MIME content type, when reported. */
  contentType?: string;
  /** Size in bytes, when reported. */
  size?: number;
  /** `true` for inline attachments (embedded images, etc.). */
  isInline?: boolean;
}

/**
 * Summary of a message as returned when listing/searching a folder
 * (`FindItem`). Cheap fields only — no body. Use {@link MailMessage} for the
 * full message.
 */
export interface MailSummary {
  /** Opaque item id used by the OWA API. */
  id: string;
  /** Change key associated with the item id. */
  changeKey?: string;
  subject: string;
  /** Sender display name + address. */
  from?: EmailAddress;
  /** Preview snippet of the body, when reported. */
  preview?: string;
  /** When the message was received (Inbox) or sent. */
  received?: Date;
  /** Whether the message has been read. */
  isRead?: boolean;
  /** Whether the message carries attachments. */
  hasAttachments?: boolean;
  /** Size in bytes, when reported. */
  size?: number;
}

/** An email address with an optional display name. */
export interface EmailAddress {
  name?: string;
  address?: string;
}

/** A fully-loaded message (`GetItem`). */
export interface MailMessage extends MailSummary {
  to: EmailAddress[];
  cc: EmailAddress[];
  bcc: EmailAddress[];
  /** Body as HTML or plain text. */
  body?: { contentType: "HTML" | "Text"; content: string };
  /** Attachment descriptors (download bytes with `downloadAttachment`). */
  attachments: Attachment[];
  /** Raw RFC-822 MIME content, when requested (base64-decoded to a string). */
  mime?: string;
  /** Internet message headers, when reported. */
  headers?: Record<string, string>;
}

/** Options for {@link OutlookScraper.listMessages}. */
export interface ListOptions {
  /** Maximum number of messages to return. Default `50`. */
  limit?: number;
  /** Offset into the folder for paging. Default `0`. */
  offset?: number;
}

/** A structured mailbox search. Fields are ANDed together. */
export interface SearchQuery {
  /** Free-text search across the message (OWA `QueryString`). */
  text?: string;
  /** Sender address or name contains. */
  from?: string;
  /** Subject contains. */
  subject?: string;
  /** Only messages received on/after this date (inclusive). */
  since?: Date;
  /** Only messages received on/before this date (inclusive). */
  until?: Date;
  /** Only messages with attachments. */
  hasAttachments?: boolean;
  /** Folder to search in (distinguished name or folder id). Default `inbox`. */
  folder?: string;
  /** Maximum number of results. Default `50`. */
  limit?: number;
}

/** Export format for {@link OutlookScraper.exportMessage}. */
export type ExportFormat = "eml" | "mbox";

/** Options for {@link OutlookScraper.exportMessage}. */
export interface ExportOptions {
  /** Destination directory. Default: current working directory. */
  destDir?: string;
  /** `"eml"` (one file per message) or `"mbox"` (append to a mailbox). Default `"eml"`. */
  format?: ExportFormat;
  /** For `mbox`: file to append to. Default `mailbox.mbox` under `destDir`. */
  mboxFile?: string;
  /** Overwrite an existing `.eml` instead of skipping it. Default `false`. */
  overwrite?: boolean;
}

/** Options for {@link OutlookScraper.downloadAttachments}. */
export interface AttachmentDownloadOptions {
  /** Destination directory. Default: current working directory. */
  destDir?: string;
  /** Include inline attachments (embedded images). Default `false`. */
  includeInline?: boolean;
  /** Overwrite files that already exist locally. Default `false` (skip). */
  overwrite?: boolean;
}

/** Options for constructing an {@link OutlookScraper}. */
export interface OutlookScraperOptions {
  /**
   * Directory used to persist the browser profile (cookies, tokens) so the user
   * does not have to re-authenticate every run. Default: `./.outlook-session`.
   */
  sessionDir?: string;
  /**
   * Mail app origin. Default `https://outlook.office.com`. Override for a
   * sovereign/other cloud (e.g. `https://outlook.office365.com`).
   */
  origin?: string;
  /** Show the browser window. Required for interactive MFA login. Default `true`. */
  headed?: boolean;
  /** Preferred strategy. Default `"rest"`. */
  strategy?: Strategy;
  /** Max time (ms) to wait for the user to complete interactive login. Default `300000`. */
  loginTimeout?: number;
  /** Browser locale. Default: the system locale (e.g. `en-US`). */
  locale?: string;
  /** Per-request/navigation timeout (ms). Default `30000`. */
  timeout?: number;
}
