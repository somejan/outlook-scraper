/**
 * Live integration test — talks to a REAL mailbox, so it is opt-in.
 *
 * Enable it by persisting a session first (`outlook-scraper login`) and running:
 *   OUTLOOK_LIVE=1 npm run test:live
 *
 * Without OUTLOOK_LIVE the whole suite is skipped so CI / `npm test` stay offline.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { OutlookScraper } from "../src/index.js";

const LIVE = process.env.OUTLOOK_LIVE === "1";

test("live: connect, list folders and inbox", { skip: !LIVE }, async () => {
  const scraper = new OutlookScraper({
    sessionDir: process.env.OUTLOOK_SESSION ?? ".outlook-session",
    headed: false,
  });
  await scraper.connect();
  try {
    const folders = await scraper.listFolders();
    assert.ok(folders.length > 0, "expected at least one folder");
    assert.ok(
      folders.some((f) => /inbox|bandeja/i.test(f.displayName)),
      "expected an inbox-like folder"
    );

    const inbox = await scraper.listMessages("inbox", { limit: 3 });
    assert.ok(Array.isArray(inbox));
    if (inbox[0]) {
      const full = await scraper.getMessage(inbox[0].id);
      assert.equal(full.id, inbox[0].id);
      assert.ok("body" in full);
    }
  } finally {
    await scraper.close();
  }
});
