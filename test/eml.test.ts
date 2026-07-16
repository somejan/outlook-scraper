import { test } from "node:test";
import assert from "node:assert/strict";

import { emlFileName, emlContent, mboxEntry } from "../src/export.js";
import type { MailMessage } from "../src/types.js";

const RAW_MIME =
  "From: ana@x.es\r\nSubject: Hola\r\n\r\nCuerpo del mensaje\r\nFrom the top\r\n";

function makeMessage(overrides: Partial<MailMessage> = {}): MailMessage {
  return {
    id: "AAMkAExampleId0123456789==",
    subject: "Acta de reunión",
    received: new Date("2026-03-04T10:00:00Z"),
    to: [],
    cc: [],
    bcc: [],
    attachments: [],
    from: { name: "Ana", address: "ana@x.es" },
    mime: RAW_MIME,
    ...overrides,
  };
}

test("emlFileName: safe, dated, includes short id suffix", () => {
  const name = emlFileName(makeMessage());
  assert.match(name, /^2026-03-04_Acta_de_reuni.n_[A-Za-z0-9]{8}\.eml$/);
  assert.ok(!name.includes("/"));
  assert.ok(!name.includes(" "));
});

test("emlFileName: falls back when subject/date missing", () => {
  const name = emlFileName(makeMessage({ subject: "", received: undefined }));
  assert.match(name, /^nodate_no-subject_/);
});

test("emlContent: returns the raw MIME", () => {
  assert.equal(emlContent(makeMessage()), RAW_MIME);
});

test("emlContent: throws when MIME is missing", () => {
  assert.throws(() => emlContent(makeMessage({ mime: undefined })), /MIME/);
});

test("mboxEntry: adds a From_ separator and >-quotes body From lines", () => {
  const entry = mboxEntry(makeMessage());
  assert.ok(entry.startsWith("From ana@x.es "));
  // The body line "From the top" must be escaped to ">From the top".
  assert.ok(entry.includes(">From the top"));
  // The real header "From: ana@x.es" must NOT be escaped.
  assert.ok(entry.includes("From: ana@x.es"));
  assert.ok(entry.endsWith("\n\n"));
});
