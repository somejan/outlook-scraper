# CLI reference

All commands share these options:

| Option | Default | Description |
| --- | --- | --- |
| `--session <dir>` | `.outlook-session` | Directory for the persisted browser session. |
| `--origin <url>` | `https://outlook.office.com` | Mail app origin (change for a different cloud). |
| `--strategy <mode>` | `rest` | `auto` \| `rest` \| `dom` (only `rest` is implemented today). |
| `--headless` | off | Run the browser headless. Only works once a session is persisted. |
| `--locale <locale>` | system locale | Browser locale. |
| `--timeout <ms>` | `30000` | Per-operation timeout. |

## `login`

Open a browser to authenticate and persist the session. Run this once; the first
run shows a visible window for sign-in (including MFA).

```bash
outlook-scraper login
```

## `folders`

List mail folders (name and unread/total counts).

```bash
outlook-scraper folders
outlook-scraper folders --json
```

## `list`

List messages in a folder, newest first.

```bash
outlook-scraper list --folder inbox --limit 20
outlook-scraper list --folder sentitems --offset 20 --limit 20
outlook-scraper list --json
```

| Option | Default | Description |
| --- | --- | --- |
| `-f, --folder <name>` | `inbox` | Distinguished name or opaque folder id. |
| `-n, --limit <n>` | `50` | Max messages. |
| `--offset <n>` | `0` | Paging offset. |
| `--json` | off | Output JSON. |

Each entry prints its message `id`, which the other commands consume.

## `read`

Read a full message (headers + body + attachment names).

```bash
outlook-scraper read --id <messageId>
outlook-scraper read --id <messageId> --json
```

## `search`

Search a folder. All provided fields are ANDed together.

```bash
outlook-scraper search --from ana@x.es --subject acta --since 2026-01-01
outlook-scraper search --text presupuesto --has-attachments --folder inbox
```

| Option | Description |
| --- | --- |
| `-f, --folder <name>` | Folder to search (default `inbox`). |
| `--text <q>` | Free-text query. |
| `--from <addr>` | Sender contains. |
| `--subject <s>` | Subject contains. |
| `--since <date>` | Received on/after `YYYY-MM-DD`. |
| `--until <date>` | Received on/before `YYYY-MM-DD`. |
| `--has-attachments` | Only messages with attachments. |
| `-n, --limit <n>` | Max results (default `50`). |

## `export`

Export a single message or a whole folder to `.eml` (one file each) or `.mbox`
(appended to one mailbox file).

```bash
outlook-scraper export --id <messageId> --format eml -d ./downloads
outlook-scraper export --folder inbox --limit 200 --format mbox -d ./downloads
```

| Option | Default | Description |
| --- | --- | --- |
| `-i, --id <id>` | — | Export a single message. |
| `-f, --folder <name>` | — | Export a whole folder (mutually exclusive with `--id`). |
| `-n, --limit <n>` | `50` | Max messages when exporting a folder. |
| `-d, --dest <dir>` | `downloads` | Destination directory. |
| `--format <fmt>` | `eml` | `eml` or `mbox`. |
| `--overwrite` | off | Overwrite existing `.eml` files. |

## `attachments`

Download a message's file attachments.

```bash
outlook-scraper attachments --id <messageId> -d ./downloads
outlook-scraper attachments --id <messageId> --include-inline
```

| Option | Default | Description |
| --- | --- | --- |
| `-i, --id <id>` | — | Message id (required). |
| `-d, --dest <dir>` | `downloads` | Destination directory. |
| `--include-inline` | off | Also download inline attachments (embedded images). |
| `--overwrite` | off | Overwrite existing files. |
