import { test } from "node:test";
import assert from "node:assert/strict";

import { buildQueryString } from "../src/actions/search.js";
import { folderIdFragment, isDistinguished } from "../src/actions/ews.js";
import { OwaClient, OwaError } from "../src/owa-client.js";

test("buildQueryString: single keyword fields", () => {
  assert.equal(buildQueryString({ from: "ana@x.es" }), "from:ana@x.es");
  assert.equal(buildQueryString({ subject: "acta" }), "subject:acta");
  assert.equal(buildQueryString({ hasAttachments: true }), "hasattachments:yes");
});

test("buildQueryString: quotes values containing whitespace", () => {
  assert.equal(buildQueryString({ subject: "acta reunión" }), 'subject:"acta reunión"');
  assert.equal(buildQueryString({ from: "Ana Pérez" }), 'from:"Ana Pérez"');
});

test("buildQueryString: date range uses received>=/<= with YYYY-MM-DD", () => {
  const q = buildQueryString({
    since: new Date("2026-01-01T00:00:00Z"),
    until: new Date("2026-06-30T00:00:00Z"),
  });
  assert.equal(q, "received>=2026-01-01 received<=2026-06-30");
});

test("buildQueryString: combines fields ANDed, free text last", () => {
  const q = buildQueryString({
    from: "ana@x.es",
    subject: "acta",
    since: new Date("2026-01-01T00:00:00Z"),
    text: "presupuesto",
  });
  assert.equal(q, "from:ana@x.es subject:acta received>=2026-01-01 presupuesto");
});

test("buildQueryString: empty query yields empty string", () => {
  assert.equal(buildQueryString({}), "");
});

test("isDistinguished recognises well-known folders case-insensitively", () => {
  assert.equal(isDistinguished("inbox"), true);
  assert.equal(isDistinguished("Inbox"), true);
  assert.equal(isDistinguished("SentItems"), true);
  assert.equal(isDistinguished("some-opaque-id=="), false);
});

test("OwaClient refuses mutating actions without touching the network", async () => {
  const client = new OwaClient({ auth: undefined } as any, 1000);
  for (const action of ["DeleteItem", "MoveItem", "EmptyFolder", "SendItem"]) {
    await assert.rejects(
      () => client.call(action, {}),
      (e) => e instanceof OwaError && /read-only/.test(e.message),
      `${action} should be refused as not read-only`
    );
  }
});

test("OwaClient lets read-only actions past the guard (then fails for lack of a token)", async () => {
  const client = new OwaClient({ auth: undefined } as any, 1000);
  // FindFolder passes the allowlist, so the error is about the missing token —
  // proving the refusal above is the guard, not an unrelated failure.
  await assert.rejects(
    () => client.call("FindFolder", {}),
    (e) => e instanceof OwaError && !/read-only/.test(e.message)
  );
});

test("folderIdFragment: distinguished vs opaque id", () => {
  assert.deepEqual(folderIdFragment("inbox"), {
    __type: "DistinguishedFolderId:#Exchange",
    Id: "inbox",
  });
  assert.deepEqual(folderIdFragment("AAMkAelvis=="), {
    __type: "FolderId:#Exchange",
    Id: "AAMkAelvis==",
  });
});
