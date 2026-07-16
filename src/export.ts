/**
 * Export messages to disk as `.eml` (one file per message) or `.mbox` (many
 * messages appended into a single Unix mbox file).
 *
 * The message bytes come from OWA as RFC-822 MIME (`GetItem` with
 * `IncludeMimeContent`), so `.eml` is just those bytes written out. `.mbox`
 * wraps each message in the standard `From ` separator line with `>`-quoting of
 * body lines that would otherwise look like a separator.
 */

import type { MailMessage } from "./types.js";

/** Filesystem-unsafe characters and whitespace runs, collapsed to `_`. */
const UNSAFE = /[\\/:*?"<>|]+|\s+/g;

/** Build a safe `.eml` file name from a message's subject + date. */
export function emlFileName(message: MailMessage): string {
  const date = message.received
    ? message.received.toISOString().slice(0, 10)
    : "nodate";
  const subject = (message.subject || "no-subject")
    .replace(UNSAFE, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);
  const shortId = message.id.replace(/[^A-Za-z0-9]/g, "").slice(-8);
  return `${date}_${subject}_${shortId}.eml`;
}

/** The RFC-822 bytes for a message. Requires the message was loaded with MIME. */
export function emlContent(message: MailMessage): string {
  if (!message.mime) {
    throw new Error(
      "Message has no MIME content — load it with includeMime=true before exporting."
    );
  }
  return message.mime;
}

/**
 * Build one mbox entry for a message: the `From ` separator line, the MIME with
 * `From ` lines in the body quoted as `>From `, and a trailing blank line.
 */
export function mboxEntry(message: MailMessage): string {
  const mime = emlContent(message);
  const sender = message.from?.address ?? "MAILER-DAEMON";
  const when = (message.received ?? new Date()).toUTCString();
  const quoted = mime.replace(/^From /gm, ">From ");
  return `From ${sender} ${when}\n${quoted}\n\n`;
}
