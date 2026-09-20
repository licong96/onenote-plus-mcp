import { afterEach, describe, expect, it } from 'vitest';
import {
  CLIENT_ID_ENV,
  TENANT_ID_ENV,
  TOKEN_CACHE_ENV,
  getAuthority,
  getClientId,
  getConfigDir,
  getTokenCachePath,
  getTokenCacheSeed,
} from '@/config.js';
import { homedir } from 'node:os';
import { join } from 'node:path';

describe('getConfigDir', () => {
  const savedXdg = process.env.XDG_CONFIG_HOME;

  afterEach(() => {
    if (savedXdg === undefined) delete process.env.XDG_CONFIG_HOME;
    else process.env.XDG_CONFIG_HOME = savedXdg;
  });

  it('defaults to ~/.config/onenote-mcp when XDG_CONFIG_HOME is unset', () => {
    delete process.env.XDG_CONFIG_HOME;
    expect(getConfigDir()).toBe(join(homedir(), '.config', 'onenote-mcp'));
  });

  it('uses XDG_CONFIG_HOME when set', () => {
    process.env.XDG_CONFIG_HOME = '/custom/config';
    expect(getConfigDir()).toBe(join('/custom/config', 'onenote-mcp'));
  });

  it('falls back to default when XDG_CONFIG_HOME is an empty string', () => {
    process.env.XDG_CONFIG_HOME = '';
    expect(getConfigDir()).toBe(join(homedir(), '.config', 'onenote-mcp'));
  });
});

describe('getTokenCachePath', () => {
  const savedXdg = process.env.XDG_CONFIG_HOME;

  afterEach(() => {
    if (savedXdg === undefined) delete process.env.XDG_CONFIG_HOME;
    else process.env.XDG_CONFIG_HOME = savedXdg;
  });

  it('appends tokens.json to the config dir', () => {
    process.env.XDG_CONFIG_HOME = '/xdg';
    expect(getTokenCachePath()).toBe(join('/xdg', 'onenote-mcp', 'tokens.json'));
  });
});

describe('getClientId', () => {
  const saved = process.env[CLIENT_ID_ENV];

  afterEach(() => {
    if (saved === undefined) delete process.env[CLIENT_ID_ENV];
    else process.env[CLIENT_ID_ENV] = saved;
  });

  it('throws when the env var is unset', () => {
    delete process.env[CLIENT_ID_ENV];
    expect(() => getClientId()).toThrow(/ONENOTE_MCP_CLIENT_ID/);
  });

  it('throws when the env var is an empty string', () => {
    process.env[CLIENT_ID_ENV] = '';
    expect(() => getClientId()).toThrow(/ONENOTE_MCP_CLIENT_ID/);
  });

  it('throws when the env var is whitespace only', () => {
    process.env[CLIENT_ID_ENV] = '   ';
    expect(() => getClientId()).toThrow(/ONENOTE_MCP_CLIENT_ID/);
  });

  it('returns the trimmed value when set', () => {
    process.env[CLIENT_ID_ENV] = '  abc-123  ';
    expect(getClientId()).toBe('abc-123');
  });
});

describe('getAuthority', () => {
  const saved = process.env[TENANT_ID_ENV];

  afterEach(() => {
    if (saved === undefined) delete process.env[TENANT_ID_ENV];
    else process.env[TENANT_ID_ENV] = saved;
  });

  it('returns the /common authority when TENANT_ID is unset', () => {
    delete process.env[TENANT_ID_ENV];
    expect(getAuthority()).toBe('https://login.microsoftonline.com/common');
  });

  it('returns a tenant-specific authority when TENANT_ID is set', () => {
    process.env[TENANT_ID_ENV] = 'my-tenant-id';
    expect(getAuthority()).toBe('https://login.microsoftonline.com/my-tenant-id');
  });

  it('returns the /common authority when TENANT_ID is empty', () => {
    process.env[TENANT_ID_ENV] = '';
    expect(getAuthority()).toBe('https://login.microsoftonline.com/common');
  });

  it('returns the /common authority when TENANT_ID is whitespace only', () => {
    process.env[TENANT_ID_ENV] = '   ';
    expect(getAuthority()).toBe('https://login.microsoftonline.com/common');
  });

  it('trims whitespace from the tenant ID', () => {
    process.env[TENANT_ID_ENV] = '  tenant-xyz  ';
    expect(getAuthority()).toBe('https://login.microsoftonline.com/tenant-xyz');
  });
});

describe('getTokenCacheSeed', () => {
  const saved = process.env[TOKEN_CACHE_ENV];

  afterEach(() => {
    if (saved === undefined) delete process.env[TOKEN_CACHE_ENV];
    else process.env[TOKEN_CACHE_ENV] = saved;
  });

  it('returns undefined when the env var is unset', () => {
    delete process.env[TOKEN_CACHE_ENV];
    expect(getTokenCacheSeed()).toBeUndefined();
  });

  it('returns undefined when the env var is empty', () => {
    process.env[TOKEN_CACHE_ENV] = '';
    expect(getTokenCacheSeed()).toBeUndefined();
  });

  it('returns undefined when the env var is whitespace only', () => {
    process.env[TOKEN_CACHE_ENV] = '   ';
    expect(getTokenCacheSeed()).toBeUndefined();
  });

  it('returns the trimmed value when set', () => {
    process.env[TOKEN_CACHE_ENV] = '  {"token":"abc"}  ';
    expect(getTokenCacheSeed()).toBe('{"token":"abc"}');
  });
});
