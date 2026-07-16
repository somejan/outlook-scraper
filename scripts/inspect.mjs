/**
 * Diagnostic: reuse the persisted session, open the mailbox, and dump the real
 * host + cookies so we can see what OWA actually sets for this tenant.
 *
 * Run AFTER a login (so .outlook-session exists) and with no other
 * outlook-scraper process holding the session open:
 *
 *   node scripts/inspect.mjs
 *
 * Set HEADED=1 to watch the window:  HEADED=1 node scripts/inspect.mjs
 */

import { chromium } from "playwright";

const sessionDir = process.env.OUTLOOK_SESSION ?? ".outlook-session";
const origin = process.env.OUTLOOK_ORIGIN ?? "https://outlook.office.com";
const headed = process.env.HEADED === "1";

const context = await chromium.launchPersistentContext(sessionDir, {
  headless: !headed,
  viewport: headed ? null : { width: 1600, height: 1000 },
});

const page = context.pages()[0] ?? (await context.newPage());
await page.goto(`${origin}/mail/`, { waitUntil: "domcontentloaded" });
// Give redirects a moment to settle.
await page.waitForTimeout(5000);

console.log("Final URL after loading the mailbox:");
console.log("  " + page.url());

const cookies = await context.cookies();
console.log(`\n${cookies.length} cookies. Names + domains:`);
for (const c of cookies) {
  const flag = /canary|owa|token|sccauth/i.test(c.name) ? "  <-- interesting" : "";
  console.log(`  ${c.name}  @ ${c.domain}${flag}`);
}

const canary = cookies.find((c) => c.name === "X-OWA-CANARY");
console.log(
  "\nX-OWA-CANARY present:",
  canary ? `yes (domain ${canary.domain})` : "NO"
);

await context.close();
