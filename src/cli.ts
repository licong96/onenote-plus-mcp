#!/usr/bin/env node
import { login, logout } from './auth/index.js';
import { seedTokenCacheFromEnv } from './auth/tokenCache.js';
import { runServer, SERVER_VERSION } from './index.js';
import { getHttpHost, getHttpPort, getHttpToken } from './config.js';
import { startHttpServer } from './http/server.js';

const HELP = `onenote-mcp v${SERVER_VERSION}

Usage:
  onenote-mcp                          Run the MCP server over stdio (default).
  onenote-mcp --transport http         Run the MCP server over HTTP at /mcp.
                                       Requires ONENOTE_MCP_HTTP_TOKEN.
  onenote-mcp login                    Sign in via Microsoft device-code flow
                                       and cache the refresh token at
                                       ~/.config/onenote-mcp/tokens.json.
  onenote-mcp logout                   Remove cached tokens and sign out.
  onenote-mcp --help                   Show this help.
  onenote-mcp --version                Print the version.

HTTP transport flags:
  --transport <stdio|http>             Transport to use. Default: stdio.
  --host <addr>                        Bind address. Default: 127.0.0.1.
  --port <n>                           Listen port. Default: 3000.

Environment:
  ONENOTE_MCP_CLIENT_ID                (required) Application (client) ID of
                                       your Microsoft Entra app registration.
  ONENOTE_MCP_TENANT_ID                (optional) Tenant ID; defaults to
                                       /common which works for personal + work
                                       accounts.
  ONENOTE_MCP_HTTP_TOKEN               (required for --transport http) Shared
                                       bearer token clients must send as
                                       \`Authorization: Bearer <token>\`.
  ONENOTE_MCP_HTTP_HOST                (optional) Bind address; --host wins.
  ONENOTE_MCP_HTTP_PORT                (optional) Listen port; --port wins.
  ONENOTE_MCP_TOKEN_CACHE              (optional) Verbatim tokens.json contents
                                       used to seed the cache on a headless
                                       host that can't run device-code login.
                                       Applied only in --transport http mode
                                       when no cached token exists yet.
  XDG_CONFIG_HOME                      (optional) Override the config dir.
`;

interface ParsedArgs {
  subcommand: string | undefined;
  transport: 'stdio' | 'http';
  host?: string;
  port?: number;
}

const fail = (message: string, exitCode = 1): never => {
  process.stderr.write(`error: ${message}\n`);
  process.exit(exitCode);
};

const parseTransport = (value: string | undefined): 'stdio' | 'http' => {
  if (value === 'stdio' || value === 'http') return value;
  return fail(`--transport must be \`stdio\` or \`http\` (got: ${value ?? '<missing>'})`);
};

const parseArgs = (argv: readonly string[]): ParsedArgs => {
  let subcommand: string | undefined;
  let transport: 'stdio' | 'http' = 'stdio';
  let host: string | undefined;
  let port: number | undefined;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    if (arg === '--transport') {
      transport = parseTransport(argv[i + 1]);
      i += 1;
    } else if (arg.startsWith('--transport=')) {
      transport = parseTransport(arg.slice('--transport='.length));
    } else if (arg === '--host') {
      host = argv[i + 1];
      if (host === undefined) fail('--host requires a value');
      i += 1;
    } else if (arg.startsWith('--host=')) {
      host = arg.slice('--host='.length);
    } else if (arg === '--port') {
      const value = argv[i + 1];
      if (value === undefined) fail('--port requires a value');
      port = Number(value);
      i += 1;
    } else if (arg.startsWith('--port=')) {
      port = Number(arg.slice('--port='.length));
    } else if (arg.startsWith('-')) {
      if (subcommand === undefined) subcommand = arg;
      else fail(`unexpected flag: ${arg}`);
    } else if (subcommand === undefined) {
      subcommand = arg;
    } else {
      fail(`unexpected argument: ${arg}`);
    }
  }

  return { subcommand, transport, host, port };
};

const cmdLogin = async (): Promise<void> => {
  const account = await login(({ message }) => {
    process.stdout.write(`${message}\n`);
  });
  process.stdout.write(`\nSigned in as ${account.username}.\n`);
};

const cmdLogout = async (): Promise<void> => {
  await logout();
  process.stdout.write('Signed out. Cached tokens removed.\n');
};

const cmdServe = async (args: ParsedArgs): Promise<void> => {
  if (args.transport === 'stdio') {
    await runServer();
    return;
  }
  const host = getHttpHost(args.host);
  const port = getHttpPort(args.port);
  const token = getHttpToken();
  if (await seedTokenCacheFromEnv()) {
    process.stderr.write('Seeded OneNote token cache from ONENOTE_MCP_TOKEN_CACHE.\n');
  }
  const running = await startHttpServer({ host, port, token });
  process.stderr.write(`onenote-mcp HTTP server listening on http://${host}:${running.port}\n`);
  process.stderr.write(`  POST ${running.url}  (Authorization: Bearer <token>)\n`);
  process.stderr.write(`  GET  http://${host}:${running.port}/healthz\n`);
  const shutdown = (signal: NodeJS.Signals): void => {
    process.stderr.write(`\nReceived ${signal}, shutting down.\n`);
    void running.close().then(() => process.exit(0));
  };
  process.once('SIGINT', shutdown);
  process.once('SIGTERM', shutdown);
};

const main = async (): Promise<void> => {
  const args = parseArgs(process.argv.slice(2));

  switch (args.subcommand) {
    case undefined:
    case 'serve':
      await cmdServe(args);
      return;
    case 'login':
      await cmdLogin();
      return;
    case 'logout':
      await cmdLogout();
      return;
    case '-h':
    case '--help':
    case 'help':
      process.stdout.write(HELP);
      return;
    case '-v':
    case '--version':
      process.stdout.write(`${SERVER_VERSION}\n`);
      return;
    default:
      fail(`unknown subcommand: ${args.subcommand}\n\n${HELP}`);
  }
};

main().catch((err: unknown) => {
  const message = err instanceof Error ? err.message : String(err);
  fail(message);
});
