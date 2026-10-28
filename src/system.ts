/**
 * Locale and time-zone defaults taken from the machine the tool runs on.
 *
 * OWA's request header wants a *Windows* time-zone id (e.g. "Romance Standard
 * Time"), while the OS/Node reports an IANA id (e.g. "Europe/Madrid"), so a
 * mapping generated from Unicode CLDR data (see `windows-zones.ts`) is used.
 * Unmapped zones fall back to UTC.
 */

import { IANA_TO_WINDOWS } from "./windows-zones.js";

const FALLBACK_LOCALE = "en-US";

/** The system's preferred locale (e.g. `en-US`), or `en-US` if unknown. */
export function systemLocale(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().locale || FALLBACK_LOCALE;
  } catch {
    return FALLBACK_LOCALE;
  }
}

/** The system's IANA time zone (e.g. `Europe/Madrid`), or `UTC` if unknown. */
export function systemIanaTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  } catch {
    return "UTC";
  }
}

/**
 * Convert a time-zone name to the Windows id OWA expects. A value that is
 * already a Windows id (contains "Standard Time" or is "UTC") passes through;
 * unmapped IANA zones fall back to `UTC`.
 */
export function toWindowsTimeZone(zone: string): string {
  if (zone === "UTC" || /Standard Time|Daylight Time/.test(zone)) return zone;
  return Object.hasOwn(IANA_TO_WINDOWS, zone) ? IANA_TO_WINDOWS[zone] : "UTC";
}

/** The system time zone as a Windows id. */
export function systemWindowsTimeZone(): string {
  return toWindowsTimeZone(systemIanaTimeZone());
}
