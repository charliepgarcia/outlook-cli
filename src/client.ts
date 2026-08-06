/**
 * Fetch wrapper for Microsoft Graph. Handles auth injection, the
 * Prefer/Content-Type headers, 429/503/504 backoff, and the Graph error
 * envelope. Retry policy is the safety-critical part of this file — see
 * docs/02-write-endpoints.md §8: sendMail/reply/replyAll/forward/draft-send
 * have no documented idempotency guarantee, so a blindly-retried POST can
 * double-send a real email. This client NEVER auto-retries POST.
 */

import { GRAPH_BASE, type RequestSpec } from './types.js';
import { parseGraphError } from './errors.js';

export interface GraphClientDeps {
  /** Supplies a fresh bearer token per attempt. Never imported from auth.ts here — DI only. */
  getToken: () => Promise<string>;
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  /** Default: GRAPH_BASE from types.ts. */
  baseUrl?: string;
  /**
   * Retries allowed *after* the initial attempt, for idempotent verbs only
   * (GET/DELETE/PATCH) on 429/503/504. Default 4 ⇒ up to 5 total attempts.
   */
  maxRetries?: number;
}

/** docs/02 §8: only these verbs are safe to retry on a transient failure. */
const RETRYABLE_METHODS = new Set(['GET', 'DELETE', 'PATCH']);
/** 429 (throttling, docs/00 §7) plus the two transient gateway statuses named in the brief. */
const RETRYABLE_STATUSES = new Set([429, 503, 504]);
const DEFAULT_MAX_RETRIES = 4;
const BASE_BACKOFF_MS = 500;
const MAX_BACKOFF_MS = 30_000;
const ABSOLUTE_URL_RE = /^https?:\/\//i;

export class GraphClient {
  private readonly getToken: () => Promise<string>;
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly baseUrl: string;
  private readonly maxRetries: number;

  constructor(deps: GraphClientDeps) {
    this.getToken = deps.getToken;
    this.fetchImpl = deps.fetchImpl ?? fetch;
    this.sleep = deps.sleep ?? defaultSleep;
    this.baseUrl = deps.baseUrl ?? GRAPH_BASE;
    this.maxRetries = deps.maxRetries ?? DEFAULT_MAX_RETRIES;
  }

  /**
   * Pure request-spec builder — no `getToken()` call, no network. This is
   * what `--dry-run` renders. Authorization is attached later, only inside
   * the actual send path.
   */
  buildSpec(
    method: string,
    path: string,
    body?: unknown,
    query?: Record<string, unknown>,
  ): RequestSpec {
    const upperMethod = method.toUpperCase();
    const url = this.buildUrl(path, query);
    const headers: Record<string, string> = {};

    if (upperMethod === 'GET') {
      // Keeps message/attachment ids stable across folder moves (docs/01 §8).
      headers['Prefer'] = 'IdType="ImmutableId"';
    }
    if (body !== undefined) {
      headers['Content-Type'] = 'application/json';
    }

    const spec: RequestSpec = { method: upperMethod, url, headers };
    if (body !== undefined) spec.body = body;
    return spec;
  }

  async request<T>(spec: RequestSpec): Promise<T> {
    const res = await this.executeWithRetry(spec);
    return parseSuccessBody<T>(res);
  }

  /** Raw Response, for binary reads (e.g. attachment/$value download). */
  async raw(spec: RequestSpec): Promise<Response> {
    return this.executeWithRetry(spec);
  }

  async get<T>(
    path: string,
    query?: Record<string, string | number | boolean | undefined>,
  ): Promise<T> {
    return this.request<T>(this.buildSpec('GET', path, undefined, query));
  }

  async post<T>(path: string, body?: unknown): Promise<T> {
    return this.request<T>(this.buildSpec('POST', path, body));
  }

  async patch<T>(path: string, body: unknown): Promise<T> {
    return this.request<T>(this.buildSpec('PATCH', path, body));
  }

  async del(path: string): Promise<void> {
    await this.request<void>(this.buildSpec('DELETE', path));
  }

  private async send(spec: RequestSpec): Promise<Response> {
    const token = await this.getToken();
    const headers: Record<string, string> = {
      ...(spec.headers ?? {}),
      Authorization: `Bearer ${token}`,
    };
    const init: RequestInit = { method: spec.method, headers };
    if (spec.body !== undefined) {
      init.body = JSON.stringify(spec.body);
    }
    return this.fetchImpl(spec.url, init);
  }

  private async executeWithRetry(spec: RequestSpec): Promise<Response> {
    const method = spec.method.toUpperCase();
    const retryEligible = RETRYABLE_METHODS.has(method);
    let attempt = 0;

    for (;;) {
      const res = await this.send(spec);
      if (res.ok) return res;

      const willRetry =
        retryEligible && RETRYABLE_STATUSES.has(res.status) && attempt < this.maxRetries;

      if (!willRetry) {
        const bodyText = await res.text().catch(() => '');
        const retryAfter = parseRetryAfterSeconds(res.headers.get('retry-after'));
        throw parseGraphError(res.status, bodyText, retryAfter);
      }

      const delayMs = computeDelayMs(res.headers.get('retry-after'), attempt);
      await this.sleep(delayMs);
      attempt += 1;
    }
  }

  private buildUrl(path: string, query?: Record<string, unknown>): string {
    const isAbsolute = ABSOLUTE_URL_RE.test(path);
    const base = isAbsolute ? path : joinUrl(this.baseUrl, path);

    // An absolute `path` is (by contract) a server-issued @odata.nextLink
    // that already carries its own fully-encoded query string — appending
    // more query params here would be the classic re-prefix/double-query bug.
    if (isAbsolute || !query) return base;

    const qs = buildQueryString(query);
    if (!qs) return base;
    return `${base}${base.includes('?') ? '&' : '?'}${qs}`;
  }
}

function joinUrl(base: string, path: string): string {
  const trimmedBase = base.endsWith('/') ? base.slice(0, -1) : base;
  const normalizedPath = path.startsWith('/') ? path : `/${path}`;
  return `${trimmedBase}${normalizedPath}`;
}

/**
 * OData system params ($filter, $select, $top, ...) conventionally keep a
 * literal `$` — percent-encoding it to `%24` is legal but pointlessly
 * obscures the request and mismatches every Graph doc example verbatim
 * ("must NOT be over-encoded into uselessness"). Values are always encoded.
 */
function buildQueryString(query: Record<string, unknown>): string {
  const parts: string[] = [];
  for (const [key, rawValue] of Object.entries(query)) {
    if (rawValue === undefined || rawValue === null) continue;
    parts.push(`${encodeQueryKey(key)}=${encodeURIComponent(String(rawValue))}`);
  }
  return parts.join('&');
}

function encodeQueryKey(key: string): string {
  return key.startsWith('$') ? `$${encodeURIComponent(key.slice(1))}` : encodeURIComponent(key);
}

/** Retry-After is documented in seconds (docs/00 §7); defend against an HTTP-date form too. */
function parseRetryAfterSeconds(header: string | null): number | undefined {
  if (header == null) return undefined;
  const trimmed = header.trim();
  if (trimmed.length === 0) return undefined;

  const asNumber = Number(trimmed);
  if (Number.isFinite(asNumber) && asNumber >= 0) return asNumber;

  const asDate = Date.parse(trimmed);
  if (!Number.isNaN(asDate)) {
    return Math.max(0, (asDate - Date.now()) / 1000);
  }
  return undefined;
}

function computeDelayMs(retryAfterHeader: string | null, attempt: number): number {
  const seconds = parseRetryAfterSeconds(retryAfterHeader);
  if (seconds !== undefined) return seconds * 1000;
  return fullJitterBackoffMs(attempt);
}

/** Full-jitter exponential backoff, capped ~30s, for the no-Retry-After case. */
function fullJitterBackoffMs(attempt: number): number {
  const cap = Math.min(MAX_BACKOFF_MS, BASE_BACKOFF_MS * 2 ** attempt);
  return Math.random() * cap;
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function parseSuccessBody<T>(res: Response): Promise<T> {
  if (res.status === 204) return undefined as T;
  const text = await res.text();
  if (!text) return undefined as T;
  try {
    return JSON.parse(text) as T;
  } catch {
    // A non-JSON 2xx body shouldn't happen against Graph, but never crash on it.
    return undefined as T;
  }
}
