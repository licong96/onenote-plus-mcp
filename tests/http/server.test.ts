import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HTTP_HOST_ENV,
  HTTP_PORT_ENV,
  HTTP_TOKEN_ENV,
  getHttpHost,
  getHttpPort,
  getHttpToken,
} from '@/config.js';
import { startHttpServer, type RunningHttpServer } from '@/http/server.js';

vi.mock('../../src/auth/index.js', () => ({
  getAccessToken: vi.fn(async () => 'fake-access-token'),
}));

const TOKEN = 'test-bearer-token-do-not-use';

const initializePayload = {
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2025-03-26',
    capabilities: {},
    clientInfo: { name: 'vitest', version: '0.0.0' },
  },
};

const parseMcpResponseBody = async (res: Response): Promise<unknown> => {
  const text = await res.text();
  const contentType = res.headers.get('content-type') ?? '';
  if (contentType.includes('text/event-stream')) {
    for (const line of text.split('\n')) {
      if (line.startsWith('data: ')) return JSON.parse(line.slice('data: '.length));
    }
    throw new Error(`SSE body had no data frame: ${text}`);
  }
  return JSON.parse(text);
};

const mcpHeaders = (token: string | undefined): Record<string, string> => {
  const h: Record<string, string> = {
    'Content-Type': 'application/json',
    Accept: 'application/json, text/event-stream',
  };
  if (token !== undefined) h.Authorization = `Bearer ${token}`;
  return h;
};

describe('HTTP server', () => {
  let server: RunningHttpServer;
  let base: string;

  beforeEach(async () => {
    server = await startHttpServer({ host: '127.0.0.1', port: 0, token: TOKEN });
    base = `http://127.0.0.1:${server.port}`;
  });

  afterEach(async () => {
    await server.close();
  });

  it('GET /healthz returns 200 without authentication', async () => {
    const res = await fetch(`${base}/healthz`);
    expect(res.status).toBe(200);
    expect(await res.text()).toBe('');
  });

  it('POST /mcp without Authorization returns 401', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(undefined),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
    expect(res.headers.get('www-authenticate')).toMatch(/^Bearer /);
  });

  it('POST /mcp with wrong bearer token returns 401', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders('not-the-token'),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
  });

  it('POST /mcp with wrong token of identical length still returns 401', async () => {
    const sameLength = 'a'.repeat(TOKEN.length);
    expect(sameLength.length).toBe(TOKEN.length);
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(sameLength),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
  });

  it('POST /mcp with valid token initializes and returns the server info', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as {
      jsonrpc: string;
      id: number;
      result: { serverInfo: { name: string } };
    };
    expect(body.jsonrpc).toBe('2.0');
    expect(body.id).toBe(1);
    expect(body.result.serverInfo.name).toBe('onenote-mcp');
  });

  it('POST /mcp with valid token + tools/list returns the registered OneNote tools', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/list',
        params: {},
      }),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as {
      result: { tools: { name: string }[] };
    };
    const names = body.result.tools.map((t) => t.name);
    expect(names).toContain('list_notebooks');
    expect(names).toContain('create_page');
    expect(names).toContain('delete_page');
  });

  it('unknown paths return 404', async () => {
    const res = await fetch(`${base}/nope`, {
      headers: { Authorization: `Bearer ${TOKEN}` },
    });
    expect(res.status).toBe(404);
  });
});

describe('config getters for HTTP transport', () => {
  beforeEach(() => {
    delete process.env[HTTP_TOKEN_ENV];
    delete process.env[HTTP_HOST_ENV];
    delete process.env[HTTP_PORT_ENV];
  });

  it('getHttpToken throws when ONENOTE_MCP_HTTP_TOKEN is unset', () => {
    expect(() => getHttpToken()).toThrow(/ONENOTE_MCP_HTTP_TOKEN/);
  });

  it('getHttpToken throws when ONENOTE_MCP_HTTP_TOKEN is whitespace', () => {
    process.env[HTTP_TOKEN_ENV] = '   ';
    expect(() => getHttpToken()).toThrow(/ONENOTE_MCP_HTTP_TOKEN/);
  });

  it('getHttpToken returns trimmed value', () => {
    process.env[HTTP_TOKEN_ENV] = '  secret  ';
    expect(getHttpToken()).toBe('secret');
  });

  it('getHttpHost prefers the override over env and default', () => {
    process.env[HTTP_HOST_ENV] = '10.0.0.1';
    expect(getHttpHost('192.168.1.1')).toBe('192.168.1.1');
    expect(getHttpHost()).toBe('10.0.0.1');
    delete process.env[HTTP_HOST_ENV];
    expect(getHttpHost()).toBe('127.0.0.1');
  });

  it('getHttpPort prefers the override, accepts string env, validates range', () => {
    process.env[HTTP_PORT_ENV] = '4000';
    expect(getHttpPort(8080)).toBe(8080);
    expect(getHttpPort()).toBe(4000);
    expect(getHttpPort('5000')).toBe(5000);
    expect(() => getHttpPort('99999')).toThrow(/Invalid HTTP port/);
    expect(() => getHttpPort('not-a-port')).toThrow(/Invalid HTTP port/);
    delete process.env[HTTP_PORT_ENV];
    expect(getHttpPort()).toBe(3000);
  });
});
