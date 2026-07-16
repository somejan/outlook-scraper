/**
 * outlook-scraper — read and download Office 365 / Outlook Web (OWA) email
 * (folders, messages, attachments) using a persisted, interactively
 * authenticated browser session. No Graph API registration required.
 *
 * @example
 * ```ts
 * import { OutlookScraper } from "outlook-scraper";
 *
 * const scraper = new OutlookScraper();
 * await scraper.connect();                    // opens a browser for MFA login the first time
 * const inbox = await scraper.listMessages("inbox", { limit: 20 });
 * const full  = await scraper.getMessage(inbox[0].id);
 * await scraper.exportMessage(inbox[0].id, { destDir: "./downloads", format: "eml" });
 * await scraper.close();
 * ```
 */

export { OutlookScraper, type AttachmentResult } from "./OutlookScraper.js";
// OwaClient is intentionally not exported: it is an internal detail and the
// scraper's read-only surface should not expose it. OwaError is public so
// callers can catch it.
export { OwaError } from "./owa-client.js";
export { buildQueryString } from "./actions/search.js";
export { emlFileName, emlContent, mboxEntry } from "./export.js";

export type {
  OutlookScraperOptions,
  Strategy,
  DistinguishedFolder,
  MailFolder,
  MailSummary,
  MailMessage,
  EmailAddress,
  Attachment,
  EwsId,
  ListOptions,
  SearchQuery,
  ExportFormat,
  ExportOptions,
  AttachmentDownloadOptions,
} from "./types.js";
