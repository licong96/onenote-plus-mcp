import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { AccountInfo } from '@azure/msal-node';
import { HTTP_HOST_ENV, HTTP_PORT_ENV, HTTP_TOKEN_ENV } from '@/config.js';

const VERSION = '9.9.9-test';

const mockLogin = vi.fn();
const mockLogout = vi.fn();
const mockRunServer = vi.fn();
const mockSeedTokenCache = vi.fn();
const mockStartHttpServer = vi.fn();
const mockClose = vi.fn();

vi.mock('../src/auth/index.js', () => ({
  login: (...args: unknown[]) => mockLogin(...args),
  logout: (...args: unknown[]) => mockLogout(...args),
}));

vi.mock('../src/auth/tokenCache.js', () => ({
  seedTokenCacheFromEnv: (...args: unknown[]) => mockSeedTokenCache(...args),
}));

vi.mock('../src/index.js', () => ({
  SERVER_VERSION: VERSION,
  runServer: (...args: unknown[]) => mockRunServer(...args),
}));

vi.mock('../src/http/server.js', () => ({
  startHttpServer: (...args: unknown[]) => mockStartHttpServer(...args),
}));

const fakeAccount: AccountInfo = {
  homeAccountId: 'home-id',
  environment: 'login.microsoftonline.com',
  tenantId: 'tenant-id',
  username: 'user@example.com',
  localAccountId: 'local-id',
};

const savedEnv: Record<string, string | undefined> = {};
const ENV_KEYS = [HTTP_TOKEN_ENV, HTTP_HOST_ENV, HTTP_PORT_ENV];

let stdout: string;
let stderr: string;
let exitCodes: number[];
let savedArgv: string[];
let savedSignalListeners: Record<'SIGINT' | 'SIGTERM', NodeJS.SignalsListener[]>;

// cli.ts has no exports — it runs main() on import — so each case re-imports it
// with a fresh module registry after seeding process.argv.
const runCli = async (...argv: string[]): Promise<void> => {
  process.argv = ['node', '/tmp/onenote-mcp/cli.js', ...argv];
  vi.resetModules();
  await import('@/cli.js');
  for (let i = 0; i < 3; i += 1) await new Promise((resolve) => setTimeout(resolve, 0));
};

const signalHandler = (signal: 'SIGINT' | 'SIGTERM'): NodeJS.SignalsListener => {
  const added = process
    .listeners(signal)
    .filter((fn) => !savedSignalListeners[signal].includes(fn));
  expect(added).toHaveLength(1);
  return added[0]!;
};

beforeEach(() => {
  vi.clearAllMocks();
  stdout = '';
  stderr = '';
  exitCodes = [];
  savedArgv = process.argv;
  savedSignalListeners = {
    SIGINT: process.listeners('SIGINT'),
    SIGTERM: process.listeners('SIGTERM'),
  };
  for (const key of ENV_KEYS) savedEnv[key] = process.env[key];
  process.env[HTTP_TOKEN_ENV] = 'test-http-token';
  delete process.env[HTTP_HOST_ENV];
  delete process.env[HTTP_PORT_ENV];

  vi.spyOn(process.stdout, 'write').mockImplementation((chunk) => {
    stdout += String(chunk);
    return true;
  });
  vi.spyOn(process.stderr, 'write').mockImplementation((chunk) => {
    stderr += String(chunk);
    return true;
  });
  // Non-throwing so the top-level `main().catch(fail)` never leaves a rejected promise.
  vi.spyOn(process, 'exit').mockImplementation(((code?: number) => {
    exitCodes.push(code ?? 0);
  }) as never);

  mockLogin.mockResolvedValue(fakeAccount);
  mockLogout.mockResolvedValue(undefined);
  mockRunServer.mockResolvedValue(undefined);
  mockSeedTokenCache.mockResolvedValue(false);
  mockClose.mockResolvedValue(undefined);
  mockStartHttpServer.mockImplementation(async (opts: { port: number }) => ({
    url: `http://127.0.0.1:${opts.port || 41234}/mcp`,
    port: opts.port || 41234,
    close: mockClose,
  }));
});

afterEach(() => {
  process.argv = savedArgv;
  for (const key of ENV_KEYS) {
    if (savedEnv[key] === undefined) delete process.env[key];
    else process.env[key] = savedEnv[key];
  }
  for (const signal of ['SIGINT', 'SIGTERM'] as const) {
    for (const fn of process.listeners(signal)) {
      if (!savedSignalListeners[signal].includes(fn)) process.removeListener(signal, fn);
    }
  }
  vi.restoreAllMocks();
});

describe('subcommand dispatch', () => {
  it('runs the stdio server when invoked with no arguments', async () => {
    await runCli();
    expect(mockRunServer).toHaveBeenCalledOnce();
    expect(mockStartHttpServer).not.toHaveBeenCalled();
    expect(exitCodes).toEqual([]);
  });

  it('runs the stdio server for an explicit `serve` subcommand', async () => {
    await runCli('serve');
    expect(mockRunServer).toHaveBeenCalledOnce();
  });

  it('runs the stdio server for --transport stdio', async () => {
    await runCli('serve', '--transport', 'stdio');
    expect(mockRunServer).toHaveBeenCalledOnce();
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  it('signs in and reports the account for `login`', async () => {
    await runCli('login');
    expect(mockLogin).toHaveBeenCalledOnce();
    expect(stdout).toContain('Signed in as user@example.com.');
    expect(mockRunServer).not.toHaveBeenCalled();
  });

  it('relays the device-code prompt to stdout during login', async () => {
    mockLogin.mockImplementation(async (onPrompt: (r: { message: string }) => void) => {
      onPrompt({ message: 'Go to https://microsoft.com/devicelogin and enter ABC-123' });
      return fakeAccount;
    });
    await runCli('login');
    expect(stdout).toContain('Go to https://microsoft.com/devicelogin and enter ABC-123');
  });

  it('clears cached tokens for `logout`', async () => {
    await runCli('logout');
    expect(mockLogout).toHaveBeenCalledOnce();
    expect(stdout).toContain('Signed out. Cached tokens removed.');
  });

  it.each(['-h', '--help', 'help'])('prints usage for %s', async (flag) => {
    await runCli(flag);
    expect(stdout).toContain(`onenote-mcp v${VERSION}`);
    expect(stdout).toContain('Usage:');
    expect(stdout).toContain('--transport <stdio|http>');
    expect(mockRunServer).not.toHaveBeenCalled();
    expect(exitCodes).toEqual([]);
  });

  it.each(['-v', '--version'])('prints the bare version for %s', async (flag) => {
    await runCli(flag);
    expect(stdout).toBe(`${VERSION}\n`);
    expect(mockRunServer).not.toHaveBeenCalled();
  });

  it('exits 1 with usage on an unknown subcommand', async () => {
    await runCli('frobnicate');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('unknown subcommand: frobnicate');
    expect(stderr).toContain('Usage:');
    expect(mockRunServer).not.toHaveBeenCalled();
  });
});

describe('HTTP transport arguments', () => {
  it('passes space-separated --host/--port through to the HTTP server', async () => {
    await runCli('--transport', 'http', '--host', '0.0.0.0', '--port', '8080');
    expect(mockStartHttpServer).toHaveBeenCalledWith({
      host: '0.0.0.0',
      port: 8080,
      token: 'test-http-token',
    });
    expect(mockRunServer).not.toHaveBeenCalled();
  });

  it('accepts the --flag=value form', async () => {
    await runCli('--transport=http', '--host=10.0.0.5', '--port=9001');
    expect(mockStartHttpServer).toHaveBeenCalledWith({
      host: '10.0.0.5',
      port: 9001,
      token: 'test-http-token',
    });
  });

  it('falls back to host/port env vars when the flags are absent', async () => {
    process.env[HTTP_HOST_ENV] = '192.168.1.10';
    process.env[HTTP_PORT_ENV] = '4321';
    await runCli('--transport', 'http');
    expect(mockStartHttpServer).toHaveBeenCalledWith({
      host: '192.168.1.10',
      port: 4321,
      token: 'test-http-token',
    });
  });

  it('lets flags win over host/port env vars', async () => {
    process.env[HTTP_HOST_ENV] = '192.168.1.10';
    process.env[HTTP_PORT_ENV] = '4321';
    await runCli('--transport', 'http', '--host', '127.0.0.9', '--port', '5555');
    expect(mockStartHttpServer).toHaveBeenCalledWith({
      host: '127.0.0.9',
      port: 5555,
      token: 'test-http-token',
    });
  });

  it('defaults host and port when neither flags nor env are set', async () => {
    await runCli('--transport', 'http');
    expect(mockStartHttpServer).toHaveBeenCalledWith({
      host: '127.0.0.1',
      port: 3000,
      token: 'test-http-token',
    });
  });

  it('advertises the listening endpoints on stderr', async () => {
    await runCli('--transport', 'http', '--port', '7777');
    expect(stderr).toContain('listening on http://127.0.0.1:7777');
    expect(stderr).toContain('POST http://127.0.0.1:7777/mcp');
    expect(stderr).toContain('GET  http://127.0.0.1:7777/healthz');
    expect(stdout).toBe('');
  });

  it('reports when the token cache was seeded from the environment', async () => {
    mockSeedTokenCache.mockResolvedValue(true);
    await runCli('--transport', 'http');
    expect(stderr).toContain('Seeded OneNote token cache from ONENOTE_MCP_TOKEN_CACHE.');
  });

  it('stays quiet about seeding when there was nothing to seed', async () => {
    await runCli('--transport', 'http');
    expect(mockSeedTokenCache).toHaveBeenCalledOnce();
    expect(stderr).not.toContain('Seeded OneNote token cache');
  });

  it('does not touch the token cache in stdio mode', async () => {
    await runCli();
    expect(mockSeedTokenCache).not.toHaveBeenCalled();
  });

  it.each(['SIGINT', 'SIGTERM'] as const)('closes the server and exits 0 on %s', async (signal) => {
    await runCli('--transport', 'http');
    const handler = signalHandler(signal);
    handler(signal);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(stderr).toContain(`Received ${signal}, shutting down.`);
    expect(mockClose).toHaveBeenCalledOnce();
    expect(exitCodes).toEqual([0]);
  });
});

describe('argument errors', () => {
  it('rejects an unrecognised --transport value', async () => {
    await runCli('--transport', 'carrier-pigeon');
    expect(exitCodes[0]).toBe(1);
    expect(stderr).toContain('--transport must be `stdio` or `http` (got: carrier-pigeon)');
  });

  it('reports a missing --transport value as <missing>', async () => {
    await runCli('--transport');
    expect(exitCodes[0]).toBe(1);
    expect(stderr).toContain('(got: <missing>)');
  });

  it('rejects --host with no value', async () => {
    await runCli('--host');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('--host requires a value');
  });

  it('rejects --port with no value', async () => {
    await runCli('--port');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('--port requires a value');
  });

  it('rejects a second positional argument', async () => {
    await runCli('serve', 'extra');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('unexpected argument: extra');
  });

  it('rejects a second flag-shaped subcommand', async () => {
    await runCli('--help', '--version');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('unexpected flag: --version');
  });

  it('surfaces a non-numeric --port as a config error', async () => {
    await runCli('--transport', 'http', '--port', 'eighty');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain('Invalid HTTP port');
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  it('surfaces a missing HTTP bearer token', async () => {
    delete process.env[HTTP_TOKEN_ENV];
    await runCli('--transport', 'http');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toContain(`Missing ${HTTP_TOKEN_ENV}`);
    expect(mockStartHttpServer).not.toHaveBeenCalled();
  });

  it('surfaces a rejected login as an error exit', async () => {
    mockLogin.mockRejectedValue(new Error('device code expired'));
    await runCli('login');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toBe('error: device code expired\n');
  });

  it('stringifies non-Error rejections', async () => {
    mockLogout.mockRejectedValue('cache is locked');
    await runCli('logout');
    expect(exitCodes).toEqual([1]);
    expect(stderr).toBe('error: cache is locked\n');
  });
});
