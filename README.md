# outlook-scraper

Read and download **Office 365 / Outlook Web** email — folders, messages,
attachments — using a **persisted, interactively-authenticated browser session**.
No Graph API app registration required.

It is the email counterpart of
[`onedrive-scraper`](https://github.com/codeurjc/onedrive-scraper) and follows the
same approach: a Playwright browser lets you sign in once (including MFA), the
session is saved to disk, and from then on the tool calls the **same internal API
the Outlook web app uses** (`/owa/service.svc`).

Useful when Graph API / registered-app access is blocked but you can still open
your mailbox in a browser.

> **Read-only by design.** The client only issues read actions
> (`FindFolder`, `GetFolder`, `FindItem`, `GetItem`) — any other action is
> refused in code (see [src/owa-client.ts](src/owa-client.ts)). It never deletes,
> moves, or sends mail. Use it on your own mailbox and respect your
> organisation's acceptable-use policy.

## Install

```bash
npm install github:codeurjc/outlook-scraper
npx playwright install chromium   # if the postinstall did not run it
```

## CLI

The first command opens a **visible browser** — sign in (including any MFA step).
The session is saved to `./.outlook-session` so later commands run without a new
login (and can even run `--headless`).

```bash
outlook-scraper login                                    # authenticate once
outlook-scraper folders                                  # list mail folders
outlook-scraper list --folder inbox --limit 20           # list messages
outlook-scraper read --id <messageId>                    # read a full message
outlook-scraper search --from ana@x.es --since 2026-01-01
outlook-scraper export --id <messageId> --format eml -d ./downloads
outlook-scraper export --folder inbox --format mbox -d ./downloads
outlook-scraper attachments --id <messageId> -d ./downloads
```

`--folder` accepts a distinguished name (`inbox`, `sentitems`, `drafts`,
`deleteditems`, `junkemail`, `archive`, `outbox`, `msgfolderroot`) or an opaque
folder id from `folders --json`. Message ids come from `list` / `search`.

See [docs/CLI.md](docs/CLI.md) for the full reference.

## Library

```ts
import { OutlookScraper } from "outlook-scraper";

const scraper = new OutlookScraper({ sessionDir: ".outlook-session" });
await scraper.connect();                      // opens a browser for MFA login the first time

const folders  = await scraper.listFolders();
const messages = await scraper.listMessages("inbox", { limit: 50 });
const full     = await scraper.getMessage(messages[0].id);   // body + attachment info
const hits     = await scraper.search({ from: "ana@x.es", subject: "acta", since: new Date("2026-01-01") });

await scraper.exportMessage(messages[0].id, { destDir: "./downloads", format: "eml" });
await scraper.downloadAttachments(messages[0].id, { destDir: "./downloads" });

await scraper.close();
```

See [docs/API.md](docs/API.md) for the full API.

## How it works

- **Auth / session** ([src/browser.ts](src/browser.ts)) — a Playwright *persistent
  context* keeps cookies/tokens in `sessionDir`, so MFA login is a one-time step.
  The new Outlook (`outlook.cloud.microsoft`) authenticates its API with an
  `Authorization: Bearer` token (MSAL) plus an `X-AnchorMailbox` header, not a
  cookie. That token lives in the page's JavaScript, so we **observe** the mail
  app's own authenticated `service.svc` requests (restricted to known Outlook
  hosts over HTTPS) and reuse their `Authorization` / anchor / URL. The app
  refreshes the token continuously, and we force a refresh on a 401.
- **OWA client** ([src/owa-client.ts](src/owa-client.ts)) — posts the EWS-over-JSON
  envelope to `/owa/service.svc?action=<Action>` with the observed token. An
  allowlist limits it to read-only actions.
- **Actions** ([src/actions/](src/actions/)) — `FindFolder`/`GetFolder`,
  `FindItem`/`GetItem`, and `FindItem` + `QueryString` for search.
- **Export** ([src/export.ts](src/export.ts)) — `GetItem` with `IncludeMimeContent`
  yields RFC-822 MIME, written out as `.eml` or appended to an `.mbox`.
- **Attachments** ([src/mime.ts](src/mime.ts)) — the new Outlook downloads
  attachments through a separate endpoint that needs an anti-CSRF canary we
  cannot mint, so attachments are extracted from that same message MIME instead.

## Verifying the API

The exact shapes of the `/owa/service.svc` envelopes are **not documented by
Microsoft** and can change between Outlook builds. If a call fails, capture what
the real web app sends and reconcile the payloads in `src/actions/`:

- `node scripts/inspect.mjs` — dump the mailbox host and cookie names.
- `node scripts/capture-body.mjs` — while you click around the visible window,
  print the real request bodies for `FindFolder` / `FindItem` / `GetItem` / etc.

These write `capture*.json` (which is gitignored — it contains ids and request
bodies). Review before sharing.

## Development

```bash
npm run build          # compile to dist/
npm run cli -- --help  # run the CLI with tsx (no build)
npm test               # all offline suites (query building, parsing, eml/mbox, MIME, guard)
npm run test:unit      # the pure-logic subset
npm run example        # end-to-end example (needs a real mailbox)
```

`npm test` runs offline. Live behaviour against a real mailbox is verified via the
CLI / `examples/basic.ts` (or `OUTLOOK_LIVE=1 npm run test:live`).

## Limitations

- Interactive login requires a **visible** browser the first time; headless works
  only once a session is persisted.
- Read-only: listing, reading, searching, exporting, and downloading attachments.
  No sending, moving, or deleting.
- The internal API shape may need occasional maintenance (see "Verifying the
  API").

## License

[MIT](LICENSE)
