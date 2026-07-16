/**
 * Minimal RFC-822 / MIME parser — just enough to pull attachments out of a
 * message's raw MIME (which OWA gives us via `GetItem` + `IncludeMimeContent`).
 *
 * The new Outlook downloads attachments through a separate `GetFileAttachment`
 * endpoint that needs an anti-CSRF canary we cannot mint, so instead we reuse
 * the full MIME (same source `exportMessage` already uses) and extract the
 * attachment parts here. This handles the common cases: nested multipart,
 * base64 / quoted-printable / 7bit bodies, and RFC 2231 filename encoding.
 */

export interface MimeAttachment {
  filename: string;
  contentType?: string;
  contentId?: string;
  isInline: boolean;
  content: Buffer;
}

interface MimePart {
  headers: Record<string, string>;
  body: string;
}

/** Split a raw MIME string into its headers and body (first blank line). */
function splitHeaders(raw: string): MimePart {
  const norm = raw.replace(/\r\n/g, "\n");
  const sep = norm.indexOf("\n\n");
  const headerBlock = sep === -1 ? norm : norm.slice(0, sep);
  const body = sep === -1 ? "" : norm.slice(sep + 2);
  const headers: Record<string, string> = {};
  // Unfold folded header lines (continuation lines start with whitespace).
  const unfolded = headerBlock.replace(/\n[ \t]+/g, " ");
  for (const line of unfolded.split("\n")) {
    const idx = line.indexOf(":");
    if (idx === -1) continue;
    headers[line.slice(0, idx).trim().toLowerCase()] = line.slice(idx + 1).trim();
  }
  return { headers, body };
}

/** Extract a parameter value (e.g. boundary=, filename=, name=) from a header. */
function param(header: string | undefined, name: string): string | undefined {
  if (!header) return undefined;
  // RFC 2231 extended form: name*=UTF-8''%E2%82%AC…
  const ext = new RegExp(`${name}\\*=(?:([^']*)'[^']*')?([^;]+)`, "i").exec(header);
  if (ext) {
    try {
      return decodeURIComponent(ext[2].trim().replace(/^"|"$/g, ""));
    } catch {
      /* fall through to plain form */
    }
  }
  const plain = new RegExp(`${name}=("([^"]*)"|[^;]+)`, "i").exec(header);
  if (!plain) return undefined;
  return (plain[2] ?? plain[1]).trim();
}

/** Decode a part body according to its Content-Transfer-Encoding. */
function decodeBody(body: string, encoding?: string): Buffer {
  const enc = (encoding ?? "7bit").toLowerCase();
  if (enc === "base64") {
    return Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, ""), "base64");
  }
  if (enc === "quoted-printable") {
    const decoded = body
      .replace(/=\r?\n/g, "")
      .replace(/=([A-Fa-f0-9]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
    return Buffer.from(decoded, "binary");
  }
  return Buffer.from(body, "binary");
}

/** Recursively collect attachment parts from a MIME part. */
function walk(part: MimePart, out: MimeAttachment[]): void {
  const contentType = part.headers["content-type"] ?? "";
  const disposition = part.headers["content-disposition"] ?? "";
  const isMultipart = /^multipart\//i.test(contentType);

  if (isMultipart) {
    const boundary = param(contentType, "boundary");
    if (!boundary) return;
    const marker = `--${boundary}`;
    const segments = part.body.split(marker);
    // Drop the preamble (before first boundary) and the closing epilogue.
    for (const seg of segments.slice(1)) {
      if (seg.startsWith("--")) break; // closing boundary "--boundary--"
      // The newline right after the boundary and the one right before the next
      // boundary belong to the delimiter, not the part — strip both.
      const child = splitHeaders(seg.replace(/^\n/, "").replace(/\n$/, ""));
      walk(child, out);
    }
    return;
  }

  const filename = param(disposition, "filename") ?? param(contentType, "name");
  const isAttachmentDisp = /^attachment/i.test(disposition);
  const hasContentId = !!part.headers["content-id"];
  // Treat as an attachment when it has a filename or an explicit attachment
  // disposition; skip the plain body parts (text/plain, text/html) that have
  // neither.
  if (!filename && !isAttachmentDisp) return;

  out.push({
    filename: filename ?? "attachment",
    contentType: contentType.split(";")[0].trim() || undefined,
    contentId: part.headers["content-id"]?.replace(/^<|>$/g, ""),
    isInline: /^inline/i.test(disposition) || (!isAttachmentDisp && hasContentId),
    content: decodeBody(part.body, part.headers["content-transfer-encoding"]),
  });
}

/** Parse all attachments (file + inline) out of a raw MIME message. */
export function parseMimeAttachments(mime: string): MimeAttachment[] {
  const root = splitHeaders(mime);
  const out: MimeAttachment[] = [];
  walk(root, out);
  return out;
}
