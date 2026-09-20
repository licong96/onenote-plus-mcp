import { homedir } from 'node:os';
import { join } from 'node:path';

export const SCOPES = ['Notes.ReadWrite', 'offline_access'] as const;

export const AUTHORITY = 'https://login.microsoftonline.com/common';

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

export const CLIENT_ID_ENV = 'ONENOTE_MCP_CLIENT_ID';

export const TENANT_ID_ENV = 'ONENOTE_MCP_TENANT_ID';

export const HTTP_TOKEN_ENV = 'ONENOTE_MCP_HTTP_TOKEN';

export const HTTP_HOST_ENV = 'ONENOTE_MCP_HTTP_HOST';

export const HTTP_PORT_ENV = 'ONENOTE_MCP_HTTP_PORT';

export const TOKEN_CACHE_ENV = 'ONENOTE_MCP_TOKEN_CACHE';

export const DEFAULT_HTTP_HOST = '127.0.0.1';

export const DEFAULT_HTTP_PORT = 3000;

const APP_DIR_NAME = 'onenote-mcp';

const TOKEN_CACHE_FILENAME = 'tokens.json';

export const getConfigDir = (): string => {
  const xdg = process.env.XDG_CONFIG_HOME;
  const base = xdg && xdg.length > 0 ? xdg : join(homedir(), '.config');
  return join(base, APP_DIR_NAME);
};

export const getTokenCachePath = (): string => join(getConfigDir(), TOKEN_CACHE_FILENAME);

export const getClientId = (): string => {
  const id = process.env[CLIENT_ID_ENV];
  if (!id || id.trim().length === 0) {
    throw new Error(
      `Missing ${CLIENT_ID_ENV}. Set it to the Application (client) ID of your Microsoft Entra app registration. See README for setup steps.`,
    );
  }
  return id.trim();
};

export const getAuthority = (): string => {
  const tenant = process.env[TENANT_ID_ENV]?.trim();
  if (tenant && tenant.length > 0) {
    return `https://login.microsoftonline.com/${tenant}`;
  }
  return AUTHORITY;
};

export const getHttpToken = (): string => {
  const token = process.env[HTTP_TOKEN_ENV]?.trim();
  if (!token || token.length === 0) {
    throw new Error(
      `Missing ${HTTP_TOKEN_ENV}. HTTP transport requires a bearer token — generate one with \`openssl rand -base64 32\` and set ${HTTP_TOKEN_ENV} before starting the server.`,
    );
  }
  return token;
};

export const getHttpHost = (override?: string): string => {
  const flag = override?.trim();
  if (flag && flag.length > 0) return flag;
  const env = process.env[HTTP_HOST_ENV]?.trim();
  if (env && env.length > 0) return env;
  return DEFAULT_HTTP_HOST;
};

export const getHttpPort = (override?: number | string): number => {
  const candidate = override ?? process.env[HTTP_PORT_ENV];
  if (candidate === undefined || candidate === '') return DEFAULT_HTTP_PORT;
  const parsed = typeof candidate === 'number' ? candidate : Number(candidate);
  if (!Number.isInteger(parsed) || parsed < 0 || parsed > 65535) {
    throw new Error(`Invalid HTTP port: ${String(candidate)}. Must be an integer between 0 and 65535.`);
  }
  return parsed;
};

export const getTokenCacheSeed = (): string | undefined => {
  const raw = process.env[TOKEN_CACHE_ENV]?.trim();
  return raw && raw.length > 0 ? raw : undefined;
};
