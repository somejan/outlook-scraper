/**
 * Low-level client for the new Outlook Web (OWA) internal JSON API.
 *
 * It calls the same endpoint the mail app uses:
 *   POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail
 * reusing the `Authorization: Bearer` token and `X-AnchorMailbox` header
 * observed from the app's own traffic (see {@link BrowserSession}). Requests use
 * the browser context's `APIRequestContext`, so cookies are shared too.
 *
 * The request is an "EWS over JSON" envelope. The exact shape was reconciled
 * against real captured requests (see README "Verifying the API"):
 *   { "__type": "<Action>JsonRequest:#Exchange",
 *     "Header": { "__type": "JsonRequestHeaders:#Exchange",
 *                 "RequestServerVersion": "V2018_01_08", "TimeZoneContext": {…} },
 *     "Body":   { "__type": "<Action>Request:#Exchange", … } }
 */

import type { APIRequestContext } from "playwright";
import type { BrowserSession } from "./browser.js";

/** Exchange server version advertised by the new Outlook. */
const REQUEST_SERVER_VERSION = "V2018_01_08";

/** Default time zone sent in the request header (affects date interpretation). */
const DEFAULT_TIME_ZONE = "Romance Standard Time";

/** `Prefer` header the mail app sends; keeps returned item ids immutable. */
const PREFER =
  'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"';

/**
 * Actions this client is allowed to invoke. The tool is strictly read-only, so
 * mutating actions (DeleteItem, MoveItem, EmptyFolder, SendItem, …) are refused
 * here — the single choke point for every API call, so it also protects callers
 * that use `OwaClient` directly.
 */
const READ_ONLY_ACTIONS = new Set([
  "FindFolder",
  "GetFolder",
  "FindItem",
  "GetItem",
]);

/** An OWA API error carrying the HTTP status and any server response text. */
export class OwaError extends Error {
  constructor(
    message: string,
    readonly status?: number,
    readonly responseText?: string
  ) {
    super(message);
    this.name = "OwaError";
  }
}

/** Thin client over `/owa/service.svc` using an observed Bearer token. */
export class OwaClient {
  private seq = 0;

  constructor(
    private readonly session: BrowserSession,
    private readonly timeout: number,
    private readonly timeZone: string = DEFAULT_TIME_ZONE
  ) {}

  private get request(): APIRequestContext {
    return this.session.page.context().request;
  }

  /** Build the standard EWS-JSON envelope around an action-specific body. */
  private envelope(action: string, body: Record<string, unknown>): unknown {
    return {
      __type: `${action}JsonRequest:#Exchange`,
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: REQUEST_SERVER_VERSION,
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: {
            __type: "TimeZoneDefinitionType:#Exchange",
            Id: this.timeZone,
          },
        },
      },
      Body: {
        __type: `${action}Request:#Exchange`,
        ...body,
      },
    };
  }

  /**
   * Call an OWA action and return its parsed JSON response. Retries once after
   * refreshing the token if the server rejects it as unauthorized/expired.
   */
  async call<T = any>(action: string, body: Record<string, unknown>): Promise<T> {
    if (!READ_ONLY_ACTIONS.has(action)) {
      throw new OwaError(
        `Refusing action "${action}": outlook-scraper is read-only and only permits ${[
          ...READ_ONLY_ACTIONS,
        ].join(", ")}.`
      );
    }
    try {
      return await this.callOnce<T>(action, body);
    } catch (error) {
      if (error instanceof OwaError && (error.status === 401 || error.status === 440)) {
        await this.session.refreshAuth();
        return this.callOnce<T>(action, body);
      }
      throw error;
    }
  }

  private async callOnce<T>(action: string, body: Record<string, unknown>): Promise<T> {
    const auth = this.session.auth;
    if (!auth?.authorization) {
      throw new OwaError(
        "No auth token captured yet — the mail app has not made an authenticated request. Is the session logged in?"
      );
    }
    const url = `${auth.serviceUrl}?action=${encodeURIComponent(action)}&app=Mail&n=${++this
      .seq}`;
    const headers: Record<string, string> = {
      "Content-Type": "application/json; charset=utf-8",
      Accept: "application/json, text/javascript, */*; q=0.01",
      Action: action,
      Authorization: auth.authorization,
      Prefer: PREFER,
      "X-Requested-With": "XMLHttpRequest",
    };
    if (auth.anchorMailbox) headers["X-AnchorMailbox"] = auth.anchorMailbox;

    const res = await this.request.post(url, {
      timeout: this.timeout,
      headers,
      data: this.envelope(action, body),
    });
    if (!res.ok()) {
      const text = await res.text().catch(() => "");
      throw new OwaError(
        `OWA action ${action} failed: ${res.status()} ${res.statusText()}${
          text ? " — " + text.slice(0, 1000) : ""
        }`,
        res.status(),
        text
      );
    }
    const json = (await res.json()) as any;
    assertResponseOk(action, json);
    return json as T;
  }
}

/**
 * OWA returns HTTP 200 even for logical errors: the per-action result carries a
 * `ResponseClass`/`ResponseCode`. Throw on anything that is not `Success`.
 */
function assertResponseOk(action: string, json: any): void {
  const messages: any[] =
    json?.Body?.ResponseMessages?.Items ?? (json?.Body ? [json.Body] : []);
  for (const m of messages) {
    if (m?.ResponseClass && m.ResponseClass !== "Success") {
      throw new OwaError(
        `OWA action ${action} returned ${m.ResponseClass}: ${
          m.MessageText ?? m.ResponseCode ?? "unknown error"
        }`
      );
    }
  }
}
