import { test } from "node:test";
import assert from "node:assert/strict";

import { parseMimeAttachments } from "../src/mime.js";

const pdf = Buffer.from("Hello PDF attachment").toString("base64");
const png = Buffer.from("PNGDATA").toString("base64");

const MIME = [
  "From: a@b.es",
  "Subject: Test",
  'Content-Type: multipart/mixed; boundary="OUTER"',
  "",
  "--OUTER",
  'Content-Type: multipart/alternative; boundary="INNER"',
  "",
  "--INNER",
  "Content-Type: text/plain; charset=utf-8",
  "Content-Transfer-Encoding: 7bit",
  "",
  "Hola cuerpo",
  "--INNER",
  "Content-Type: text/html; charset=utf-8",
  "",
  "<p>Hola</p>",
  "--INNER--",
  "--OUTER",
  'Content-Type: application/pdf; name="acta.pdf"',
  'Content-Disposition: attachment; filename="acta.pdf"',
  "Content-Transfer-Encoding: base64",
  "",
  pdf,
  "--OUTER",
  'Content-Type: image/png; name="firma.png"',
  'Content-Disposition: inline; filename="firma.png"',
  "Content-Transfer-Encoding: base64",
  "Content-ID: <sig123>",
  "",
  png,
  "--OUTER--",
  "",
].join("\r\n");

test("parseMimeAttachments: extracts file + inline, skips body parts", () => {
  const atts = parseMimeAttachments(MIME);
  assert.equal(atts.length, 2, "should find exactly the 2 attachment parts");

  const acta = atts.find((a) => a.filename === "acta.pdf");
  assert.ok(acta, "acta.pdf found");
  assert.equal(acta!.isInline, false);
  assert.equal(acta!.contentType, "application/pdf");
  assert.equal(acta!.content.toString(), "Hello PDF attachment");

  const firma = atts.find((a) => a.filename === "firma.png");
  assert.ok(firma, "firma.png found");
  assert.equal(firma!.isInline, true);
  assert.equal(firma!.contentId, "sig123");
  assert.equal(firma!.content.toString(), "PNGDATA");
});

test("parseMimeAttachments: quoted-printable decoding", () => {
  const mime = [
    "Content-Type: multipart/mixed; boundary=B",
    "",
    "--B",
    "Content-Type: text/plain",
    "",
    "body",
    "--B",
    "Content-Type: text/plain; name=nota.txt",
    "Content-Disposition: attachment; filename=nota.txt",
    "Content-Transfer-Encoding: quoted-printable",
    "",
    "caf=C3=A9",
    "--B--",
    "",
  ].join("\r\n");
  const atts = parseMimeAttachments(mime);
  assert.equal(atts.length, 1);
  assert.equal(atts[0].content.toString("utf8"), "café");
});

test("parseMimeAttachments: no attachments in a plain message", () => {
  const mime = ["Content-Type: text/plain", "", "just a body"].join("\r\n");
  assert.deepEqual(parseMimeAttachments(mime), []);
});
