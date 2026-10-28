import { test } from "node:test";
import assert from "node:assert/strict";

import { toWindowsTimeZone, systemLocale } from "../src/system.js";

test("toWindowsTimeZone: maps IANA ids", () => {
  assert.equal(toWindowsTimeZone("Europe/Madrid"), "Romance Standard Time");
  assert.equal(toWindowsTimeZone("America/New_York"), "Eastern Standard Time");
});

test("toWindowsTimeZone: maps zones outside the common set", () => {
  assert.equal(toWindowsTimeZone("Europe/Amsterdam"), "W. Europe Standard Time");
  assert.equal(toWindowsTimeZone("Asia/Kolkata"), "India Standard Time");
  assert.equal(toWindowsTimeZone("Europe/Kyiv"), "FLE Standard Time");
});

test("toWindowsTimeZone: does not match Object prototype keys", () => {
  assert.equal(toWindowsTimeZone("constructor"), "UTC");
});

test("toWindowsTimeZone: passes Windows ids and UTC through", () => {
  assert.equal(toWindowsTimeZone("Pacific Standard Time"), "Pacific Standard Time");
  assert.equal(toWindowsTimeZone("UTC"), "UTC");
});

test("toWindowsTimeZone: unknown zones fall back to UTC", () => {
  assert.equal(toWindowsTimeZone("Mars/Olympus_Mons"), "UTC");
});

test("systemLocale: returns a non-empty locale", () => {
  assert.ok(systemLocale().length > 0);
});
