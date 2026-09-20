import { mkdir, readFile, writeFile, chmod, stat } from 'node:fs/promises';
import { dirname } from 'node:path';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { getTokenCachePath, getTokenCacheSeed, TOKEN_CACHE_ENV } from '@/config.js';

const FILE_MODE = 0o600;
const DIR_MODE = 0o700;

const ensureSecureFile = async (path: string): Promise<void> => {
  await mkdir(dirname(path), { recursive: true, mode: DIR_MODE });
  try {
    await stat(path);
    await chmod(path, FILE_MODE);
  } catch {
    // File doesn't exist yet — it'll be created with FILE_MODE on write.
  }
};

const readCacheFile = async (path: string): Promise<string> => {
  try {
    return await readFile(path, 'utf8');
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return '';
    throw err;
  }
};

const writeCacheFile = async (path: string, contents: string): Promise<void> => {
  await mkdir(dirname(path), { recursive: true, mode: DIR_MODE });
  await writeFile(path, contents, { encoding: 'utf8', mode: FILE_MODE });
  await chmod(path, FILE_MODE);
};

// Headless deploys (Fly, containers) have no TTY for the device-code login.
// Seeding lets a tokens.json produced by `onenote-mcp login` elsewhere be
// supplied via env/secret. Skipped if the cache file already has content so a
// volume-persisted, freshly-refreshed token is never clobbered by a stale seed.
export const seedTokenCacheFromEnv = async (
  path: string = getTokenCachePath(),
): Promise<boolean> => {
  const seed = getTokenCacheSeed();
  if (seed === undefined) return false;
  const existing = await readCacheFile(path);
  if (existing.length > 0) return false;
  try {
    JSON.parse(seed);
  } catch {
    throw new Error(
      `${TOKEN_CACHE_ENV} is set but does not contain valid JSON. It must hold the verbatim contents of a tokens.json produced by \`onenote-mcp login\`.`,
    );
  }
  await writeCacheFile(path, seed);
  return true;
};

export const createFileCachePlugin = (
  path: string = getTokenCachePath(),
): ICachePlugin => ({
  beforeCacheAccess: async (context: TokenCacheContext): Promise<void> => {
    await ensureSecureFile(path);
    const data = await readCacheFile(path);
    if (data.length > 0) {
      context.tokenCache.deserialize(data);
    }
  },
  afterCacheAccess: async (context: TokenCacheContext): Promise<void> => {
    if (context.cacheHasChanged) {
      await writeCacheFile(path, context.tokenCache.serialize());
    }
  },
});
