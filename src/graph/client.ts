import { setTimeout as sleep } from 'node:timers/promises';
import { GRAPH_BASE_URL } from '@/config.js';
import { getAccessToken } from '@/auth/index.js';

export class GraphError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly code?: string,
    public readonly requestId?: string,
  ) {
    super(message);
    this.name = 'GraphError';
  }
}

export interface RequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE' | 'PUT';
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
  body?: string | Uint8Array | FormData | null;
  /** Override the default JSON Accept header. */
  accept?: string;
  /** Parse mode for the response body. */
  parse?: 'json' | 'text' | 'none';
}

const MAX_RETRIES = 4;
const BASE_BACKOFF_MS = 500;

const buildUrl = (path: string, query?: RequestOptions['query']): string => {
  const url = path.startsWith('http')
    ? new URL(path)
    : new URL(`${GRAPH_BASE_URL}${path.startsWith('/') ? path : `/${path}`}`);
  if (query) {
    for (const [k, v] of Object.entries(query)) {
      if (v === undefined) continue;
      url.searchParams.set(k, String(v));
    }
  }
  return url.toString();
};

const isRetryable = (status: number): boolean =>
  status === 429 || status === 503 || status === 504 || status === 502;

const backoffDelay = (attempt: number, retryAfterHeader: string | null): number => {
  if (retryAfterHeader) {
    const retryAfterSec = Number(retryAfterHeader);
    if (Number.isFinite(retryAfterSec) && retryAfterSec >= 0) {
      return Math.min(retryAfterSec * 1000, 30_000);
    }
  }
  // Exponential backoff with full jitter, capped at 8s.
  const exp = Math.min(BASE_BACKOFF_MS * 2 ** attempt, 8_000);
  return Math.floor(Math.random() * exp);
};

const shapeError = async (response: Response): Promise<GraphError> => {
  const requestId = response.headers.get('request-id') ?? undefined;
  let code: string | undefined;
  let message = `Graph request failed with ${response.status} ${response.statusText}`;
  try {
    const text = await response.text();
    if (text.length > 0) {
      try {
        const json = JSON.parse(text) as {
          error?: { code?: string; message?: string };
        };
        if (json.error) {
          code = json.error.code;
          if (json.error.message) message = json.error.message;
        } else {
          message = text.slice(0, 500);
        }
      } catch {
        message = text.slice(0, 500);
      }
    }
  } catch {
    // Body unreadable — keep the default message.
  }
  return new GraphError(message, response.status, code, requestId);
};

export const graphRequest = async <T = unknown>(
  path: string,
  options: RequestOptions = {},
): Promise<T> => {
  const { method = 'GET', query, headers = {}, body = null, accept, parse = 'json' } = options;
  const url = buildUrl(path, query);
  const token = await getAccessToken();

  const reqHeaders: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: accept ?? 'application/json',
    ...headers,
  };

  let lastError: GraphError | undefined;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    const response = await fetch(url, { method, headers: reqHeaders, body });

    if (response.ok) {
      if (parse === 'none' || response.status === 204) {
        // Drain body to free socket.
        await response.arrayBuffer().catch(() => undefined);
        return undefined as T;
      }
      if (parse === 'text') {
        return (await response.text()) as T;
      }
      // Read as text first so empty 200 bodies don't blow up JSON.parse.
      const text = await response.text();
      if (text.length === 0) return undefined as T;
      return JSON.parse(text) as T;
    }

    const err = await shapeError(response);
    lastError = err;

    if (!isRetryable(response.status) || attempt === MAX_RETRIES) {
      throw err;
    }

    await sleep(backoffDelay(attempt, response.headers.get('retry-after')));
  }

  throw lastError ?? new GraphError('Graph request failed after retries', 0);
};

interface PageResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

export const paginate = async <T>(
  path: string,
  options: RequestOptions = {},
): Promise<T[]> => {
  const all: T[] = [];
  let next: string | undefined = path;
  let pageOptions: RequestOptions | undefined = options;
  while (next) {
    // graphRequest yields undefined for an empty 200/204 body — stop rather than deref it.
    const page: PageResponse<T> | undefined = await graphRequest<PageResponse<T> | undefined>(
      next,
      pageOptions,
    );
    if (!page) break;
    if (Array.isArray(page.value)) all.push(...page.value);
    next = page['@odata.nextLink'];
    // nextLink is a fully-qualified URL with its own query string — clear our overrides.
    pageOptions = undefined;
  }
  return all;
};
