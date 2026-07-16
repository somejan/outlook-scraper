/**
 * Diagnostic capture: open the real Outlook in a visible window, then record the
 * internal API calls it makes while YOU interact with it. This tells us the real
 * host / endpoint / auth headers so the scraper can be pointed at them.
 *
 * Run it, then in the browser: click a folder, open a message, download an
 * attachment. After CAPTURE_SECONDS (default 60) it prints what it saw.
 *
 *   node scripts/capture.mjs
 *   CAPTURE_SECONDS=90 node scripts/capture.mjs
 *
 * It prints request METHOD + URL + a few auth-relevant headers and whether the
 * request had a body. It does NOT print header *values* for secrets (auth token,
 * cookies) — only whether they are present — so it is safe to paste back.
 */

import { chromium } from "playwright";
import { writeFile } from "node:fs/promises";

const sessionDir = process.env.OUTLOOK_SESSION ?? ".outlook-session";
const startUrl = process.env.OUTLOOK_URL ?? "https://outlook.office.com/mail/";
const seconds = Number(process.env.CAPTURE_SECONDS ?? "60");

// Endpoint URL fragments that look like the mail data API (not static assets).
const API_HINTS = [
  "service.svc",
  "/owa/",
  "/api/",
  "substrate",
  "outlook.office.com/search",
  "outlook.cloud.microsoft",
];
const IGNORE = /\.(js|css|png|jpg|jpeg|gif|svg|woff2?|ico|map)(\?|$)/i;

const context = await chromium.launchPersistentContext(sessionDir, {
  headless: false,
  viewport: null,
  args: ["--start-maximized"],
});
const page = context.pages()[0] ?? (await context.newPage());

const seen = [];
page.on("request", (req) => {
  const url = req.url();
  const type = req.resourceType();
  if (type !== "fetch" && type !== "xhr") return;
  if (IGNORE.test(url)) return;
  if (!API_HINTS.some((h) => url.includes(h))) return;

  const h = req.headers();
  let body;
  try {
    body = req.postDataJSON();
  } catch {
    body = req.postData()?.slice(0, 2000);
  }
  seen.push({
    method: req.method(),
    url,
    action: (() => {
      try {
        return new URL(url).searchParams.get("action");
      } catch {
        return null;
      }
    })(),
    authPresent: {
      authorizationBearer: !!h["authorization"],
      xOwaCanary: !!h["x-owa-canary"],
      xAnchorMailbox: h["x-anchormailbox"] ?? null,
      action: h["action"] ?? null,
      prefer: h["prefer"] ?? null,
      contentType: h["content-type"] ?? null,
    },
    hasBody: body !== undefined,
    body,
  });
});

console.log(`Opening ${startUrl} …`);
await page.goto(startUrl, { waitUntil: "domcontentloaded" }).catch(() => {});
console.log(
  `\n>>> Now, in the browser window: click a folder, open a message, and download an attachment.\n>>> Capturing for ${seconds}s…\n`
);
await page.waitForTimeout(seconds * 1000);

// De-duplicate by (method + endpoint path + action) to keep the summary short.
const byKey = new Map();
for (const r of seen) {
  const path = (() => {
    try {
      const u = new URL(r.url);
      return u.host + u.pathname;
    } catch {
      return r.url;
    }
  })();
  const key = `${r.method} ${path} action=${r.action ?? ""}`;
  if (!byKey.has(key)) byKey.set(key, r);
}

console.log(`Captured ${seen.length} API request(s), ${byKey.size} unique:\n`);
for (const [key, r] of byKey) {
  console.log("• " + key);
  console.log("    auth:", JSON.stringify(r.authPresent));
}

await writeFile(
  "capture.json",
  JSON.stringify([...byKey.values()], null, 2),
  "utf8"
);
console.log(
  "\nFull detail (with request bodies) written to capture.json — review it before sharing."
);

await context.close();
