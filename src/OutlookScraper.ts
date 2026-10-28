/**
 * High-level Outlook Web (OWA) email scraper.
 *
 * Presents a small, task-focused API over a mailbox using a persisted
 * authenticated browser session:
 *   - `listFolders()`               — enumerate mail folders
 *   - `listMessages(folder)`        — list messages in a folder
 *   - `getMessage(id)`              — read a full message (body + attachments)
 *   - `search(query)`              — structured mailbox search
 *   - `exportMessage(id)`           — save a message as .eml / .mbox
 *   - `downloadAttachments(id)`     — save a message's attachments
 *
 * No Graph API app registration is required — it reuses the same internal
 * endpoint the Outlook web UI calls.
 */

import { mkdir, writeFile, appendFile, access } from "node:fs/promises";
import { constants as FS } from "node:fs";
import { join } from "node:path";

import { BrowserSession } from "./browser.js";
import { OwaClient } from "./owa-client.js";
import { findFolders, getFolder } from "./actions/folders.js";
import { findItems, getItem } from "./actions/items.js";
import { searchItems } from "./actions/search.js";
import { systemLocale, systemWindowsTimeZone } from "./system.js";
import { parseMimeAttachments } from "./mime.js";
import { emlContent, emlFileName, mboxEntry } from "./export.js";
import type {
  AttachmentDownloadOptions,
  ExportOptions,
  ListOptions,
  MailFolder,
  MailMessage,
  MailSummary,
  OutlookScraperOptions,
  SearchQuery,
  Strategy,
} from "./types.js";

const DEFAULTS = {
  sessionDir: ".outlook-session",
  origin: "https://outlook.cloud.microsoft",
  headed: true,
  strategy: "rest" as Strategy,
  loginTimeout: 300_000,
  timeout: 30_000,
};

/** Result of a single attachment download. */
export interface AttachmentResult {
  name: string;
  localPath: string;
  status: "downloaded" | "skipped" | "error";
  bytes?: number;
  error?: Error;
}

export class OutlookScraper {
  private readonly session: BrowserSession;
  private readonly origin: string;
  private readonly timeout: number;
  private client?: OwaClient;
  private connected = false;

  constructor(private readonly options: OutlookScraperOptions = {}) {
    this.origin = options.origin ?? DEFAULTS.origin;
    this.timeout = options.timeout ?? DEFAULTS.timeout;
    this.session = new BrowserSession({
      sessionDir: options.sessionDir ?? DEFAULTS.sessionDir,
      origin: this.origin,
      mailUrl: `${this.origin}/mail/`,
      headed: options.headed ?? DEFAULTS.headed,
      locale: options.locale ?? systemLocale(),
      timeout: this.timeout,
      loginTimeout: options.loginTimeout ?? DEFAULTS.loginTimeout,
    });
  }

  /**
   * Launch the browser and ensure the session is authenticated (driving an
   * interactive login the first time). Must be called before any operation.
   */
  async connect(): Promise<void> {
    if (this.connected) return;
    await this.session.ensureAuthenticated();
    this.client = new OwaClient(this.session, this.timeout, systemWindowsTimeZone());
    this.connected = true;
  }

  /** Close the browser, flushing the persisted profile to disk. */
  async close(): Promise<void> {
    await this.session.close();
    this.connected = false;
  }

  // ---------------------------------------------------------------------------
  // Folders
  // ---------------------------------------------------------------------------

  /** List mail folders (Inbox, Sent Items, …) under the mailbox root. */
  async listFolders(parent = "msgfolderroot"): Promise<MailFolder[]> {
    this.assertConnected();
    return findFolders(this.client!, parent);
  }

  /** Get one folder's metadata by distinguished name or id. */
  async getFolder(folder: string): Promise<MailFolder> {
    this.assertConnected();
    return getFolder(this.client!, folder);
  }

  // ---------------------------------------------------------------------------
  // Messages
  // ---------------------------------------------------------------------------

  /** List messages in a folder (default `inbox`), newest first. */
  async listMessages(
    folder = "inbox",
    options: ListOptions = {}
  ): Promise<MailSummary[]> {
    this.assertConnected();
    return findItems(this.client!, folder, options);
  }

  /** Load a full message by id (recipients, body, attachment descriptors). */
  async getMessage(
    id: string,
    options: { changeKey?: string; includeMime?: boolean } = {}
  ): Promise<MailMessage> {
    this.assertConnected();
    return getItem(this.client!, id, options.changeKey, options.includeMime ?? false);
  }

  /** Structured mailbox search. Fields are ANDed together. */
  async search(query: SearchQuery): Promise<MailSummary[]> {
    this.assertConnected();
    return searchItems(this.client!, query);
  }

  // ---------------------------------------------------------------------------
  // Export & attachments
  // ---------------------------------------------------------------------------

  /**
   * Export a message to disk as `.eml` (default) or append it to an `.mbox`.
   * Returns the path written (or, for mbox, the mbox file path).
   */
  async exportMessage(
    id: string,
    options: ExportOptions = {}
  ): Promise<string> {
    this.assertConnected();
    const destDir = options.destDir ?? process.cwd();
    const format = options.format ?? "eml";
    const message = await this.getMessage(id, { includeMime: true });
    await mkdir(destDir, { recursive: true });

    if (format === "mbox") {
      const mboxPath = options.mboxFile ?? join(destDir, "mailbox.mbox");
      await appendFile(mboxPath, mboxEntry(message), "utf8");
      return mboxPath;
    }

    const localPath = join(destDir, emlFileName(message));
    if (!options.overwrite && (await fileExists(localPath))) {
      return localPath;
    }
    await writeFile(localPath, emlContent(message), "utf8");
    return localPath;
  }

  /** Download all of a message's file attachments to a directory. */
  async downloadAttachments(
    id: string,
    options: AttachmentDownloadOptions = {}
  ): Promise<AttachmentResult[]> {
    this.assertConnected();
    const destDir = options.destDir ?? process.cwd();
    const overwrite = options.overwrite ?? false;
    const includeInline = options.includeInline ?? false;

    // The new Outlook downloads attachments through a separate endpoint that
    // needs an anti-CSRF canary we cannot mint. Instead we reuse the message's
    // raw MIME (same source `exportMessage` uses) and extract the parts from it.
    const message = await this.getMessage(id, { includeMime: true });
    if (!message.mime) {
      throw new Error("Could not load message MIME to extract attachments.");
    }
    const wanted = parseMimeAttachments(message.mime).filter(
      (a) => includeInline || !a.isInline
    );
    await mkdir(destDir, { recursive: true });

    const results: AttachmentResult[] = [];
    const usedNames = new Set<string>();
    for (const att of wanted) {
      const localPath = join(destDir, uniqueName(safeName(att.filename), usedNames));
      const result: AttachmentResult = { name: att.filename, localPath, status: "downloaded" };
      try {
        if (!overwrite && (await fileExists(localPath))) {
          result.status = "skipped";
        } else {
          await writeFile(localPath, att.content);
          result.bytes = att.content.length;
        }
      } catch (error) {
        result.status = "error";
        result.error = error instanceof Error ? error : new Error(String(error));
      }
      results.push(result);
    }
    return results;
  }

  private assertConnected(): void {
    if (!this.connected) {
      throw new Error("Not connected. Call connect() before using the scraper.");
    }
  }
}

/** Replace path separators and awkward characters in an attachment file name. */
function safeName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "_").replace(/^\.+/, "_").trim() || "attachment";
}

/** Disambiguate repeated file names (e.g. two `image.png`) with a suffix. */
function uniqueName(name: string, used: Set<string>): string {
  if (!used.has(name)) {
    used.add(name);
    return name;
  }
  const dot = name.lastIndexOf(".");
  const base = dot === -1 ? name : name.slice(0, dot);
  const ext = dot === -1 ? "" : name.slice(dot);
  let i = 2;
  let candidate = `${base}_${i}${ext}`;
  while (used.has(candidate)) candidate = `${base}_${++i}${ext}`;
  used.add(candidate);
  return candidate;
}

async function fileExists(path: string): Promise<boolean> {
  try {
    await access(path, FS.F_OK);
    return true;
  } catch {
    return false;
  }
}
