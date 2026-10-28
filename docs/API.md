# Library API

```ts
import { OutlookScraper } from "outlook-scraper";
```

## `new OutlookScraper(options?)`

| Option | Default | Description |
| --- | --- | --- |
| `sessionDir` | `.outlook-session` | Persisted browser profile directory. |
| `origin` | `https://outlook.office.com` | Mail app origin. |
| `headed` | `true` | Show the browser window (required for interactive login). |
| `strategy` | `rest` | `auto` \| `rest` \| `dom` (only `rest` implemented). |
| `loginTimeout` | `300000` | Max ms to wait for interactive login. |
| `locale` | system locale | Browser locale. |
| `timeout` | `30000` | Per-request timeout. |

## Lifecycle

- `connect(): Promise<void>` — launch the browser and authenticate (interactive
  login the first time). Call before any operation.
- `close(): Promise<void>` — close the browser, flushing the session to disk.
The scraper deliberately does not expose the Playwright `Page` or the low-level
`OwaClient`, to keep the read-only surface small.

## Folders

- `listFolders(parent?): Promise<MailFolder[]>` — child folders under `parent`
  (default `msgfolderroot`).
- `getFolder(folder): Promise<MailFolder>` — one folder by distinguished name or
  id.

## Messages

- `listMessages(folder?, { limit?, offset? }): Promise<MailSummary[]>` — messages
  newest-first (default folder `inbox`, limit `50`).
- `getMessage(id, { changeKey?, includeMime? }): Promise<MailMessage>` — the full
  message. Pass `includeMime: true` to include RFC-822 bytes (used by export).
- `search(query): Promise<MailSummary[]>` — structured search; fields are ANDed.

```ts
interface SearchQuery {
  text?: string;          // free text
  from?: string;          // sender contains
  subject?: string;       // subject contains
  since?: Date;           // received >=
  until?: Date;           // received <=
  hasAttachments?: boolean;
  folder?: string;        // default "inbox"
  limit?: number;         // default 50
}
```

## Export & attachments

- `exportMessage(id, { destDir?, format?, mboxFile?, overwrite? }): Promise<string>`
  — write a message as `.eml` (default) or append to an `.mbox`. Returns the path
  written.
- `downloadAttachments(id, { destDir?, includeInline?, overwrite? }): Promise<AttachmentResult[]>`
  — save a message's attachments, extracted from its RFC-822 MIME
  ([src/mime.ts](../src/mime.ts)). Inline parts (embedded images) are skipped
  unless `includeInline` is set.

## Types

`MailFolder`, `MailSummary`, `MailMessage`, `EmailAddress`, `Attachment`,
`SearchQuery`, `ExportOptions`, `AttachmentDownloadOptions` — see
[src/types.ts](../src/types.ts).

## Verifying the API

Because the internal API shape is undocumented and can change, use the diagnostic
scripts to record what the real web app sends and reconcile the payloads in
`src/actions/*`:

- `node scripts/inspect.mjs` — mailbox host + cookie names.
- `node scripts/capture-body.mjs` — real request bodies for `FindFolder` /
  `FindItem` / `GetItem` / etc. while you click around the visible window.

Both write `capture*.json` (gitignored — contains ids and request bodies).
