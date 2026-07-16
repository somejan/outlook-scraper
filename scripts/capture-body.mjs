/**
 * Focused capture: print the RAW request body the mail app sends for a few key
 * service.svc actions, so we can copy their exact envelope.
 *
 * Run it, then in the browser: click into the folder list, open different
 * folders, and let the message list load. It prints each matching body.
 *
 *   node scripts/capture-body.mjs
 *
 * Bodies may contain folder/item ids (not message content) — review before
 * sharing. Output is also written to capture-bodies.json.
 */

import { chromium } from "playwright";
import { writeFile } from "node:fs/promises";

const sessionDir = process.env.OUTLOOK_SESSION ?? ".outlook-session";
const startUrl = process.env.OUTLOOK_URL ?? "https://outlook.cloud.microsoft/mail/";
const seconds = Number(process.env.CAPTURE_SECONDS ?? "60");
const WANT = new Set([
  "FindFolder",
  "FindItem",
  "GetFolder",
  "FindConversation",
  "GetAttachment",
  "GetFileAttachment",
  "GetItem",
]);

const context = await chromium.launchPersistentContext(sessionDir, {
  headless: false,
  viewport: null,
  args: ["--start-maximized"],
});
const page = context.pages()[0] ?? (await context.newPage());

const grabbed = [];
page.on("request", (req) => {
  const url = req.url();

  // Attachment downloads go to attachments.office.net (a GET, not service.svc).
  // Log the URL shape (path + which query params exist) without their values.
  if (/attachment/i.test(url) && !url.includes("/owa/service.svc")) {
    try {
      const u = new URL(url);
      const params = [...u.searchParams.keys()];
      const key = "download " + u.host + u.pathname;
      if (!grabbed.some((g) => g.action === key)) {
        grabbed.push({ action: key, body: { method: req.method(), params } });
        console.log(`\n===== ATTACHMENT DOWNLOAD (${req.method()}) =====`);
        console.log(u.host + u.pathname);
        console.log("query params:", params.join(", "));
      }
    } catch {}
    return;
  }

  if (!url.includes("/owa/service.svc")) return;
  let action = null;
  try {
    action = new URL(url).searchParams.get("action");
  } catch {}
  if (!action || !WANT.has(action)) return;

  let raw = req.postData();
  if (raw == null) {
    const buf = req.postDataBuffer();
    raw = buf ? buf.toString("utf8") : null;
  }
  if (raw == null) return;
  if (grabbed.some((g) => g.action === action)) return; // first of each action
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch {
    parsed = raw;
  }
  grabbed.push({ action, body: parsed });
  console.log(`\n===== ${action} =====`);
  console.log(JSON.stringify(parsed, null, 1));
});

console.log(`Opening ${startUrl} …`);
await page.goto(startUrl, { waitUntil: "domcontentloaded" }).catch(() => {});
console.log(
  `\n>>> In the browser: click folders in the left list and open a couple of them.\n>>> Capturing ${[...WANT].join(", ")} for ${seconds}s…\n`
);
await page.waitForTimeout(seconds * 1000);

await writeFile("capture-bodies.json", JSON.stringify(grabbed, null, 2), "utf8");
console.log(`\nWrote ${grabbed.length} bodies to capture-bodies.json`);
await context.close();
