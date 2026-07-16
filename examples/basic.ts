/**
 * Minimal end-to-end example. Run with:
 *   npm run example
 *
 * The first run opens a visible browser — sign in (including MFA). The session
 * is persisted to ./.outlook-session so later runs skip the login.
 */

import { OutlookScraper } from "../src/index.js";

async function main() {
  const scraper = new OutlookScraper({ sessionDir: ".outlook-session" });
  await scraper.connect();
  try {
    console.log("Folders:");
    for (const f of await scraper.listFolders()) {
      console.log(`  ${f.displayName} (${f.unreadCount ?? 0}/${f.totalCount ?? "?"})`);
    }

    console.log("\nLatest inbox messages:");
    const inbox = await scraper.listMessages("inbox", { limit: 5 });
    for (const m of inbox) {
      console.log(`  ${m.received?.toISOString().slice(0, 10)}  ${m.subject}`);
    }

    if (inbox[0]) {
      const full = await scraper.getMessage(inbox[0].id);
      console.log(`\nFirst message body (${full.body?.contentType}):`);
      console.log(full.body?.content.slice(0, 400));

      const path = await scraper.exportMessage(inbox[0].id, { destDir: "downloads" });
      console.log(`\nExported to ${path}`);
    }
  } finally {
    await scraper.close();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
