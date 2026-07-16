/**
 * Mailbox search via OWA `FindItem` with a `QueryString` (Advanced Query
 * Syntax — the same syntax the Outlook web search box accepts).
 *
 * We translate the structured {@link SearchQuery} into an AQS string so that
 * building the query is pure and unit-testable (see `test/restriction.test.ts`),
 * independent of the network.
 */

import type { OwaClient } from "../owa-client.js";
import type { MailSummary, SearchQuery } from "../types.js";
import { findItemResults, folderIdFragment, parseSummary } from "./ews.js";

/** Format a Date as the `YYYY-MM-DD` AQS date literal. */
function aqsDate(d: Date): string {
  return d.toISOString().slice(0, 10);
}

/** Quote an AQS value if it contains whitespace. */
function aqsValue(value: string): string {
  return /\s/.test(value) ? `"${value.replace(/"/g, '\\"')}"` : value;
}

/**
 * Build an Advanced Query Syntax string from a structured query. Exported for
 * unit testing.
 *
 * @example
 * buildQueryString({ from: "ana@x.es", subject: "acta", since: new Date("2026-01-01") })
 * // => 'from:ana@x.es subject:acta received>=2026-01-01'
 */
export function buildQueryString(query: SearchQuery): string {
  const parts: string[] = [];
  if (query.from) parts.push(`from:${aqsValue(query.from)}`);
  if (query.subject) parts.push(`subject:${aqsValue(query.subject)}`);
  if (query.hasAttachments) parts.push("hasattachments:yes");
  if (query.since) parts.push(`received>=${aqsDate(query.since)}`);
  if (query.until) parts.push(`received<=${aqsDate(query.until)}`);
  // Free text goes last so it does not get parsed as a keyword value.
  if (query.text) parts.push(aqsValue(query.text));
  return parts.join(" ");
}

/** Run a structured search and return matching message summaries. */
export async function searchItems(
  client: OwaClient,
  query: SearchQuery
): Promise<MailSummary[]> {
  const folder = query.folder ?? "inbox";
  const limit = query.limit ?? 50;
  const queryString = buildQueryString(query);
  if (!queryString) {
    throw new Error("Empty search: provide at least one of text/from/subject/since/until/hasAttachments.");
  }

  const json = await client.call("FindItem", {
    ItemShape: {
      __type: "ItemResponseShape:#Exchange",
      BaseShape: "IdOnly",
      AdditionalProperties: [
        "item:Subject",
        "item:DateTimeReceived",
        "item:Size",
        "item:HasAttachments",
        "message:From",
        "message:IsRead",
      ].map((FieldURI) => ({ __type: "PropertyUri:#Exchange", FieldURI })),
    },
    ParentFolderIds: [folderIdFragment(folder)],
    Traversal: "Shallow",
    QueryString: {
      __type: "QueryString:#Exchange",
      Value: queryString,
    },
    Paging: {
      __type: "IndexedPageView:#Exchange",
      BasePoint: "Beginning",
      MaxEntriesReturned: limit,
      Offset: 0,
    },
  });

  return findItemResults(json).map(parseSummary);
}
