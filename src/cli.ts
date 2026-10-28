#!/usr/bin/env node
/**
 * Command-line interface for outlook-scraper.
 *
 * Run `outlook-scraper --help` (or `npm run cli -- --help` during dev) for
 * usage. See docs/CLI.md for the full reference.
 */

import { Command, Option } from "commander";

import { OutlookScraper } from "./OutlookScraper.js";
import type { OutlookScraperOptions, SearchQuery, Strategy } from "./types.js";

interface CommonOpts {
  session?: string;
  origin?: string;
  strategy?: Strategy;
  headless?: boolean;
  locale?: string;
  timeout?: string;
}

function scraperOptions(opts: CommonOpts): OutlookScraperOptions {
  return {
    sessionDir: opts.session,
    origin: opts.origin,
    strategy: opts.strategy,
    headed: opts.headless ? false : true,
    locale: opts.locale,
    timeout: opts.timeout ? Number(opts.timeout) : undefined,
  };
}

/** Attach the session options shared by every command. */
function withCommon(cmd: Command): Command {
  return cmd
    .option("--session <dir>", "directory for the persisted browser session", ".outlook-session")
    .option("--origin <url>", "mail app origin", "https://outlook.cloud.microsoft")
    .addOption(
      new Option("--strategy <mode>", "access strategy").choices(["auto", "rest", "dom"]).default("rest")
    )
    .option("--headless", "run the browser headless (only works with an already-persisted session)", false)
    .option("--locale <locale>", "browser locale (default: system locale)")
    .option("--timeout <ms>", "per-operation timeout in ms", "30000");
}

async function withScraper<T>(
  common: CommonOpts,
  fn: (scraper: OutlookScraper) => Promise<T>
): Promise<T> {
  const scraper = new OutlookScraper(scraperOptions(common));
  await scraper.connect();
  try {
    return await fn(scraper);
  } finally {
    await scraper.close();
  }
}

function fmtDate(d?: Date): string {
  return d ? d.toISOString().slice(0, 16).replace("T", " ") : "";
}

function fmtAddress(a?: { name?: string; address?: string }): string {
  if (!a) return "";
  return a.name ? `${a.name} <${a.address ?? ""}>` : a.address ?? "";
}

const program = new Command();

program
  .name("outlook-scraper")
  .description(
    "Read and download Office 365 / Outlook Web email via a persisted authenticated browser session."
  )
  .version("0.1.0");

// login -----------------------------------------------------------------------
withCommon(
  program.command("login").description("Open a browser to authenticate and persist the session")
).action(async (opts: CommonOpts) => {
  await withScraper(opts, async () => {
    console.error("Authenticated. Session persisted at:", opts.session);
  });
});

// folders ---------------------------------------------------------------------
withCommon(
  program
    .command("folders")
    .description("List mail folders")
    .option("--json", "output JSON", false)
).action(async (opts: CommonOpts & { json?: boolean }) => {
  await withScraper(opts, async (scraper) => {
    const folders = await scraper.listFolders();
    if (opts.json) {
      console.log(JSON.stringify(folders, null, 2));
      return;
    }
    for (const f of folders) {
      const counts =
        f.totalCount !== undefined ? `  (${f.unreadCount ?? 0}/${f.totalCount})` : "";
      console.log(`${f.displayName}${counts}`);
    }
  });
});

// list ------------------------------------------------------------------------
withCommon(
  program
    .command("list")
    .description("List messages in a folder (default: inbox)")
    .option("-f, --folder <name>", "folder distinguished name or id", "inbox")
    .option("-n, --limit <n>", "max messages", "50")
    .option("--offset <n>", "paging offset", "0")
    .option("--json", "output JSON", false)
).action(
  async (
    opts: CommonOpts & { folder: string; limit: string; offset: string; json?: boolean }
  ) => {
    await withScraper(opts, async (scraper) => {
      const messages = await scraper.listMessages(opts.folder, {
        limit: Number(opts.limit),
        offset: Number(opts.offset),
      });
      if (opts.json) {
        console.log(
          JSON.stringify(
            messages.map((m) => ({ ...m, received: m.received?.toISOString() })),
            null,
            2
          )
        );
        return;
      }
      for (const m of messages) {
        const flag = m.isRead === false ? "●" : " ";
        const att = m.hasAttachments ? "📎" : " ";
        console.log(
          `${flag}${att} ${fmtDate(m.received).padEnd(17)} ${fmtAddress(m.from).slice(0, 28).padEnd(28)} ${m.subject}`
        );
        console.log(`     id: ${m.id}`);
      }
    });
  }
);

// read ------------------------------------------------------------------------
withCommon(
  program
    .command("read")
    .description("Read a full message by id")
    .requiredOption("-i, --id <id>", "message id (from `list`)")
    .option("--json", "output JSON", false)
).action(async (opts: CommonOpts & { id: string; json?: boolean }) => {
  await withScraper(opts, async (scraper) => {
    const m = await scraper.getMessage(opts.id);
    if (opts.json) {
      console.log(JSON.stringify({ ...m, received: m.received?.toISOString() }, null, 2));
      return;
    }
    console.log(`Subject: ${m.subject}`);
    console.log(`From:    ${fmtAddress(m.from)}`);
    console.log(`To:      ${m.to.map(fmtAddress).join(", ")}`);
    if (m.cc.length) console.log(`Cc:      ${m.cc.map(fmtAddress).join(", ")}`);
    console.log(`Date:    ${fmtDate(m.received)}`);
    if (m.attachments.length) {
      console.log(`Attachments: ${m.attachments.map((a) => a.name).join(", ")}`);
    }
    console.log("");
    console.log(m.body?.content ?? "(no body)");
  });
});

// search ----------------------------------------------------------------------
withCommon(
  program
    .command("search")
    .description("Search the mailbox")
    .option("-f, --folder <name>", "folder to search", "inbox")
    .option("--text <q>", "free-text query")
    .option("--from <addr>", "sender contains")
    .option("--subject <s>", "subject contains")
    .option("--since <date>", "received on/after (YYYY-MM-DD)")
    .option("--until <date>", "received on/before (YYYY-MM-DD)")
    .option("--has-attachments", "only messages with attachments", false)
    .option("-n, --limit <n>", "max results", "50")
    .option("--json", "output JSON", false)
).action(
  async (
    opts: CommonOpts & {
      folder: string;
      text?: string;
      from?: string;
      subject?: string;
      since?: string;
      until?: string;
      hasAttachments?: boolean;
      limit: string;
      json?: boolean;
    }
  ) => {
    await withScraper(opts, async (scraper) => {
      const query: SearchQuery = {
        folder: opts.folder,
        text: opts.text,
        from: opts.from,
        subject: opts.subject,
        since: opts.since ? new Date(opts.since) : undefined,
        until: opts.until ? new Date(opts.until) : undefined,
        hasAttachments: opts.hasAttachments,
        limit: Number(opts.limit),
      };
      const messages = await scraper.search(query);
      if (opts.json) {
        console.log(
          JSON.stringify(
            messages.map((m) => ({ ...m, received: m.received?.toISOString() })),
            null,
            2
          )
        );
        return;
      }
      for (const m of messages) {
        console.log(
          `${fmtDate(m.received).padEnd(17)} ${fmtAddress(m.from).slice(0, 28).padEnd(28)} ${m.subject}`
        );
        console.log(`     id: ${m.id}`);
      }
      console.error(`\n${messages.length} result(s).`);
    });
  }
);

// export ----------------------------------------------------------------------
withCommon(
  program
    .command("export")
    .description("Export message(s) to .eml or .mbox")
    .option("-i, --id <id>", "export a single message by id")
    .option("-f, --folder <name>", "export a whole folder")
    .option("-n, --limit <n>", "max messages when exporting a folder", "50")
    .option("-d, --dest <dir>", "destination directory", "downloads")
    .addOption(new Option("--format <fmt>", "output format").choices(["eml", "mbox"]).default("eml"))
    .option("--overwrite", "overwrite existing .eml files", false)
).action(
  async (
    opts: CommonOpts & {
      id?: string;
      folder?: string;
      limit: string;
      dest: string;
      format: "eml" | "mbox";
      overwrite?: boolean;
    }
  ) => {
    if (!opts.id && !opts.folder) {
      throw new Error("Provide --id <messageId> or --folder <name>.");
    }
    await withScraper(opts, async (scraper) => {
      const ids: string[] = opts.id
        ? [opts.id]
        : (await scraper.listMessages(opts.folder!, { limit: Number(opts.limit) })).map(
            (m) => m.id
          );
      let ok = 0;
      for (const id of ids) {
        const path = await scraper.exportMessage(id, {
          destDir: opts.dest,
          format: opts.format,
          overwrite: opts.overwrite,
        });
        ok++;
        console.error(`[${ok}/${ids.length}] → ${path}`);
      }
      console.error(`Done: ${ok} message(s) exported to ${opts.dest}`);
    });
  }
);

// attachments -----------------------------------------------------------------
withCommon(
  program
    .command("attachments")
    .description("Download a message's attachments")
    .requiredOption("-i, --id <id>", "message id")
    .option("-d, --dest <dir>", "destination directory", "downloads")
    .option("--include-inline", "also download inline attachments", false)
    .option("--overwrite", "overwrite existing files", false)
).action(
  async (
    opts: CommonOpts & {
      id: string;
      dest: string;
      includeInline?: boolean;
      overwrite?: boolean;
    }
  ) => {
    await withScraper(opts, async (scraper) => {
      const results = await scraper.downloadAttachments(opts.id, {
        destDir: opts.dest,
        includeInline: opts.includeInline,
        overwrite: opts.overwrite,
      });
      for (const r of results) {
        const tag = r.status === "downloaded" ? "↓" : r.status === "skipped" ? "=" : "✗";
        const extra =
          r.status === "downloaded" && r.bytes !== undefined
            ? ` (${r.bytes} bytes)`
            : r.status === "error"
              ? ` — ${r.error?.message}`
              : "";
        console.error(`${tag} ${r.name}${extra}`);
      }
      const failed = results.filter((r) => r.status === "error").length;
      if (failed > 0) process.exitCode = 1;
    });
  }
);

program.parseAsync(process.argv).catch((error: unknown) => {
  console.error("Error:", error instanceof Error ? error.message : error);
  process.exit(1);
});
