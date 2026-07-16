import { test } from "node:test";
import assert from "node:assert/strict";

import {
  parseAddress,
  parseAddressList,
  parseAttachments,
  parseDate,
  parseSummary,
  findItemResults,
  getItemResult,
} from "../src/actions/ews.js";

test("parseAddress: name + address from a Mailbox node", () => {
  assert.deepEqual(
    parseAddress({ Mailbox: { Name: "Ana", EmailAddress: "ana@x.es" } }),
    { name: "Ana", address: "ana@x.es" }
  );
});

test("parseAddress: tolerates a bare mailbox and missing fields", () => {
  assert.deepEqual(parseAddress({ EmailAddress: "b@x.es" }), { address: "b@x.es" });
  assert.equal(parseAddress(undefined), undefined);
  assert.equal(parseAddress({}), undefined);
});

test("parseAddressList: maps and drops empties", () => {
  const list = parseAddressList([
    { Mailbox: { Name: "A", EmailAddress: "a@x.es" } },
    { Mailbox: {} },
    { Mailbox: { EmailAddress: "c@x.es" } },
  ]);
  assert.deepEqual(list, [{ name: "A", address: "a@x.es" }, { address: "c@x.es" }]);
});

test("parseDate: parses valid, rejects junk", () => {
  assert.ok(parseDate("2026-03-04T10:00:00Z") instanceof Date);
  assert.equal(parseDate("not a date"), undefined);
  assert.equal(parseDate(undefined), undefined);
});

test("parseSummary: pulls the cheap fields off a Message item", () => {
  const summary = parseSummary({
    ItemId: { Id: "ID1", ChangeKey: "CK1" },
    Subject: "Hola",
    From: { Mailbox: { Name: "Ana", EmailAddress: "ana@x.es" } },
    DateTimeReceived: "2026-03-04T10:00:00Z",
    IsRead: false,
    HasAttachments: true,
    Size: 2048,
  });
  assert.equal(summary.id, "ID1");
  assert.equal(summary.changeKey, "CK1");
  assert.equal(summary.subject, "Hola");
  assert.deepEqual(summary.from, { name: "Ana", address: "ana@x.es" });
  assert.equal(summary.isRead, false);
  assert.equal(summary.hasAttachments, true);
  assert.equal(summary.size, 2048);
});

test("parseSummary: defaults subject when missing", () => {
  assert.equal(parseSummary({ ItemId: { Id: "X" } }).subject, "(no subject)");
});

test("parseAttachments: keeps file attachments with ids", () => {
  const atts = parseAttachments({
    Attachments: [
      { AttachmentId: { Id: "A1" }, Name: "doc.pdf", ContentType: "application/pdf", Size: 10, IsInline: false },
      { Name: "no-id-skipped" },
      { AttachmentId: { Id: "A2" }, Name: "img.png", IsInline: true },
    ],
  });
  assert.equal(atts.length, 2);
  assert.deepEqual(atts[0], {
    id: "A1",
    name: "doc.pdf",
    contentType: "application/pdf",
    size: 10,
    isInline: false,
  });
  assert.equal(atts[1].isInline, true);
});

test("findItemResults / getItemResult: dig through the response envelope", () => {
  const findResp = {
    Body: { ResponseMessages: { Items: [{ RootFolder: { Items: [{ ItemId: { Id: "A" } }] } }] } },
  };
  assert.equal(findItemResults(findResp).length, 1);
  assert.deepEqual(findItemResults({}), []);

  const getResp = {
    Body: { ResponseMessages: { Items: [{ Items: [{ ItemId: { Id: "B" } }] }] } },
  };
  assert.equal(getItemResult(getResp).ItemId.Id, "B");
});
