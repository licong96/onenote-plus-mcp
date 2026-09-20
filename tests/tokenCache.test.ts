import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtemp, rm, stat, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileCachePlugin, seedTokenCacheFromEnv } from '@/auth/tokenCache.js';
import { TOKEN_CACHE_ENV } from '@/config.js';

const makeContext = (initial: string) => {
  let serialized = initial;
  return {
    cacheHasChanged: false,
    tokenCache: {
      deserialize: (data: string): void => {
        serialized = data;
      },
      serialize: (): string => serialized,
    },
    setSerialized: (value: string): void => {
      serialized = value;
    },
    getSerialized: (): string => serialized,
  };
};

describe('createFileCachePlugin', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'onenote-mcp-test-'));
    path = join(dir, 'nested', 'tokens.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('writes the cache file with mode 600', async () => {
    const plugin = createFileCachePlugin(path);
    const ctx = makeContext('{"hello":"world"}');
    ctx.cacheHasChanged = true;
    await plugin.afterCacheAccess(ctx as never);

    const stats = await stat(path);
    // Mask off file-type bits.
    expect(stats.mode & 0o777).toBe(0o600);
    expect(await readFile(path, 'utf8')).toBe('{"hello":"world"}');
  });

  it('round-trips through beforeCacheAccess', async () => {
    const plugin = createFileCachePlugin(path);
    const writeCtx = makeContext('{"token":"abc"}');
    writeCtx.cacheHasChanged = true;
    await plugin.afterCacheAccess(writeCtx as never);

    const readCtx = makeContext('');
    await plugin.beforeCacheAccess(readCtx as never);
    expect(readCtx.getSerialized()).toBe('{"token":"abc"}');
  });

  it('treats a missing cache file as empty', async () => {
    const plugin = createFileCachePlugin(path);
    const ctx = makeContext('preset');
    await plugin.beforeCacheAccess(ctx as never);
    // No deserialize call when file is missing — preset value remains.
    expect(ctx.getSerialized()).toBe('preset');
  });

  it('does not write when cacheHasChanged is false', async () => {
    const plugin = createFileCachePlugin(path);
    const ctx = makeContext('{"x":1}');
    ctx.cacheHasChanged = false;
    await plugin.afterCacheAccess(ctx as never);
    await expect(stat(path)).rejects.toMatchObject({ code: 'ENOENT' });
  });
});

describe('seedTokenCacheFromEnv', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'onenote-mcp-seed-'));
    path = join(dir, 'nested', 'tokens.json');
    delete process.env[TOKEN_CACHE_ENV];
  });

  afterEach(async () => {
    delete process.env[TOKEN_CACHE_ENV];
    await rm(dir, { recursive: true, force: true });
  });

  it('is a no-op when the env var is unset', async () => {
    expect(await seedTokenCacheFromEnv(path)).toBe(false);
    await expect(stat(path)).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('writes the seed with mode 600 when no cache file exists', async () => {
    process.env[TOKEN_CACHE_ENV] = '{"seeded":true}';
    expect(await seedTokenCacheFromEnv(path)).toBe(true);

    const stats = await stat(path);
    expect(stats.mode & 0o777).toBe(0o600);
    expect(await readFile(path, 'utf8')).toBe('{"seeded":true}');
  });

  it('does not overwrite an existing non-empty cache file', async () => {
    const plugin = createFileCachePlugin(path);
    const ctx = makeContext('{"refreshed":true}');
    ctx.cacheHasChanged = true;
    await plugin.afterCacheAccess(ctx as never);

    process.env[TOKEN_CACHE_ENV] = '{"seeded":true}';
    expect(await seedTokenCacheFromEnv(path)).toBe(false);
    expect(await readFile(path, 'utf8')).toBe('{"refreshed":true}');
  });

  it('throws when the seed is not valid JSON', async () => {
    process.env[TOKEN_CACHE_ENV] = 'not-json';
    await expect(seedTokenCacheFromEnv(path)).rejects.toThrow(TOKEN_CACHE_ENV);
    await expect(stat(path)).rejects.toMatchObject({ code: 'ENOENT' });
  });
});
