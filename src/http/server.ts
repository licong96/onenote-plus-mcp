import { createServer as createHttpListener, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import { Buffer } from 'node:buffer';
import { timingSafeEqual } from 'node:crypto';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { createServer as createMcpServer } from '@/index.js';

export interface HttpServerOptions {
  host: string;
  port: number;
  token: string;
}

export interface RunningHttpServer {
  url: string;
  port: number;
  close: () => Promise<void>;
}

const MCP_PATH = '/mcp';
const HEALTHZ_PATH = '/healthz';

const bearerEqual = (header: string | undefined, token: string): boolean => {
  if (!header) return false;
  const expected = `Bearer ${token}`;
  if (header.length !== expected.length) return false;
  const a = Buffer.from(header);
  const b = Buffer.from(expected);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
};

const sendJson = (res: ServerResponse, status: number, body: unknown): void => {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(payload).toString(),
  });
  res.end(payload);
};

const sendUnauthorized = (res: ServerResponse): void => {
  res.writeHead(401, {
    'Content-Type': 'application/json',
    'WWW-Authenticate': 'Bearer realm="onenote-mcp"',
  });
  res.end(JSON.stringify({ error: 'unauthorized' }));
};

const handleMcpRequest = async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
  const server = createMcpServer();
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });

  let cleaned = false;
  const cleanup = async (): Promise<void> => {
    if (cleaned) return;
    cleaned = true;
    await transport.close();
    await server.close();
  };

  res.on('close', () => { void cleanup(); });

  try {
    await server.connect(transport);
    await transport.handleRequest(req, res);
  } finally {
    await cleanup();
  }
};

const route = (token: string) =>
  async (req: IncomingMessage, res: ServerResponse): Promise<void> => {
    const url = req.url ?? '/';
    const pathname = url.split('?', 1)[0] ?? '/';

    if (req.method === 'GET' && pathname === HEALTHZ_PATH) {
      res.writeHead(200);
      res.end();
      return;
    }

    if (pathname === MCP_PATH) {
      if (!bearerEqual(req.headers.authorization, token)) {
        sendUnauthorized(res);
        return;
      }
      await handleMcpRequest(req, res);
      return;
    }

    sendJson(res, 404, { error: 'not_found' });
  };

export const createHttpServer = ({ token }: HttpServerOptions): Server => {
  if (!token || token.length === 0) {
    throw new Error('HTTP transport requires a non-empty bearer token.');
  }
  const handler = route(token);
  return createHttpListener((req, res) => {
    handler(req, res).catch((err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      if (!res.headersSent) sendJson(res, 500, { error: 'internal_error', message });
      else res.end();
    });
  });
};

export const startHttpServer = async (options: HttpServerOptions): Promise<RunningHttpServer> => {
  const server = createHttpServer(options);
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(options.port, options.host, () => {
      server.off('error', reject);
      resolve();
    });
  });
  const address = server.address();
  const boundPort = typeof address === 'object' && address ? address.port : options.port;
  return {
    url: `http://${options.host}:${boundPort}${MCP_PATH}`,
    port: boundPort,
    close: () =>
      new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      }),
  };
};
