/**
 * Persistent browser session management for the new Outlook Web
 * (outlook.cloud.microsoft).
 *
 * Uses a Playwright *persistent context* so cookies and tokens survive between
 * runs — the user only completes the interactive (MFA) login once. The visible
 * browser window is what lets the user satisfy multi-factor prompts.
 *
 * The new Outlook authenticates its internal `service.svc` API with an
 * `Authorization: Bearer <token>` header (MSAL) plus an `X-AnchorMailbox`
 * header — NOT the old `X-OWA-CANARY` cookie. That token lives in the page's
 * JavaScript, not in a cookie, so we cannot mint it ourselves. Instead we
 * *observe* the mail app's own authenticated requests and reuse their
 * `Authorization` / `X-AnchorMailbox` / service URL. The app refreshes the
 * token continuously (background polling), so a captured token stays fresh, and
 * we can force a refresh by reloading the page.
 */

import { chromium, type BrowserContext, type Page } from "playwright";

/** Hosts whose `service.svc` requests we trust for observing the auth token. */
const ALLOWED_HOSTS = [
  "outlook.cloud.microsoft",
  "outlook.office.com",
  "outlook.office365.com",
];

/** True when a hostname is one of the trusted Outlook hosts (or a subdomain). */
function isAllowedHost(hostname: string): boolean {
  const host = hostname.toLowerCase();
  return ALLOWED_HOSTS.some((h) => host === h || host.endsWith(`.${h}`));
}

export interface BrowserSessionOptions {
  /** Directory holding the persisted browser profile. */
  sessionDir: string;
  /** Origin of the mail app, e.g. `https://outlook.cloud.microsoft`. */
  origin: string;
  /** URL to load when checking/establishing authentication (the mailbox). */
  mailUrl: string;
  /** Show the window. Must be `true` for interactive login. */
  headed: boolean;
  /** Browser locale. */
  locale: string;
  /** Default navigation/operation timeout (ms). */
  timeout: number;
  /** Max time (ms) to wait for interactive login to complete. */
  loginTimeout: number;
  /** Optional logger; defaults to `console.error`. */
  log?: (message: string) => void;
}

/** Auth context observed from the mail app's own `service.svc` requests. */
export interface AuthContext {
  /** `Bearer <token>` value. */
  authorization: string;
  /** `X-AnchorMailbox` header value (`PUID:…@tenant`). */
  anchorMailbox?: string;
  /** Base service URL, e.g. `https://outlook.cloud.microsoft/owa/service.svc`. */
  serviceUrl: string;
}

/** Manages the persistent Playwright context, login, and token observation. */
export class BrowserSession {
  private context?: BrowserContext;
  private _page?: Page;
  private captured?: AuthContext;
  private readonly log: (message: string) => void;

  constructor(private readonly options: BrowserSessionOptions) {
    this.log = options.log ?? ((m) => console.error(m));
  }

  /** The active page. Throws if the session has not been launched. */
  get page(): Page {
    if (!this._page) {
      throw new Error("Browser session not launched. Call launch() first.");
    }
    return this._page;
  }

  /** Whether the session has an active page. */
  get isOpen(): boolean {
    return !!this._page;
  }

  /** The most recently observed auth context, if any. */
  get auth(): AuthContext | undefined {
    return this.captured;
  }

  /** Launch the persistent context, open a page, and start observing tokens. */
  async launch(): Promise<void> {
    if (this.context) return;
    this.context = await chromium.launchPersistentContext(this.options.sessionDir, {
      headless: !this.options.headed,
      locale: this.options.locale,
      acceptDownloads: true,
      viewport: this.options.headed ? null : { width: 1600, height: 1000 },
      args: this.options.headed ? ["--start-maximized"] : [],
    });
    this.context.setDefaultTimeout(this.options.timeout);
    this.context.setDefaultNavigationTimeout(this.options.timeout);

    // Observe the mail app's own authenticated service.svc requests and reuse
    // their Authorization / anchor / URL for our API calls. Only trust HTTPS
    // requests to a known Outlook host, so the observed token can never be bound
    // to an attacker-controlled endpoint.
    this.context.on("request", (req) => {
      const url = req.url();
      if (!url.includes("/owa/service.svc")) return;
      let parsed: URL;
      try {
        parsed = new URL(url);
      } catch {
        return;
      }
      if (parsed.protocol !== "https:" || !isAllowedHost(parsed.hostname)) return;
      const headers = req.headers();
      const authorization = headers["authorization"];
      if (!authorization) return;
      this.captured = {
        authorization,
        anchorMailbox: headers["x-anchormailbox"],
        serviceUrl: url.split("?")[0],
      };
    });

    this._page = this.context.pages()[0] ?? (await this.context.newPage());
  }

  /** Navigate to the mailbox if we are not already on the mail origin. */
  async ensureOnOrigin(): Promise<void> {
    const page = this.page;
    if (!page.url().startsWith(this.options.origin)) {
      await page.goto(this.options.mailUrl, { waitUntil: "domcontentloaded" });
    }
  }

  /** Poll until an auth token has been observed, or the deadline passes. */
  private async waitForToken(timeoutMs: number): Promise<boolean> {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      if (this.captured?.authorization) return true;
      await this.page.waitForTimeout(1000);
    }
    return !!this.captured?.authorization;
  }

  /**
   * Ensure the session is authenticated and a token has been captured, driving
   * an interactive login if needed. Loading the mailbox fires a burst of
   * authenticated service.svc calls, which is how we obtain a token — so this
   * single path covers both "reuse persisted session" and "log in now".
   */
  async ensureAuthenticated(): Promise<void> {
    await this.launch();
    await this.page
      .goto(this.options.mailUrl, { waitUntil: "domcontentloaded" })
      .catch(() => {});

    // Reused session: the app loads and issues authenticated calls quickly.
    if (await this.waitForToken(20_000)) {
      this.log("[outlook-scraper] Reusing persisted authenticated session.");
      return;
    }

    if (!this.options.headed) {
      throw new Error(
        "Not authenticated and running headless. Re-run with a visible browser (headed: true) to log in interactively."
      );
    }

    this.log(
      "[outlook-scraper] Authentication required. A browser window is open — " +
        "please sign in (including any multi-factor step). Waiting…"
    );
    if (await this.waitForToken(this.options.loginTimeout)) {
      this.log("[outlook-scraper] Login successful. Session persisted.");
      return;
    }
    throw new Error(
      `Interactive login did not complete within ${Math.round(
        this.options.loginTimeout / 1000
      )}s.`
    );
  }

  /**
   * Force a fresh token by reloading the mailbox (used after a 401). Returns
   * true if a token was observed within the timeout.
   */
  async refreshAuth(timeoutMs = 30_000): Promise<boolean> {
    try {
      await this.page.goto(this.options.mailUrl, { waitUntil: "domcontentloaded" });
    } catch {
      /* ignore navigation hiccups; the observer may still fire */
    }
    return this.waitForToken(timeoutMs);
  }

  /** Close the context, flushing the persisted profile to disk. */
  async close(): Promise<void> {
    await this.context?.close();
    this.context = undefined;
    this._page = undefined;
    this.captured = undefined;
  }
}
