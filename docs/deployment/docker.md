# Deploy with Docker (VPS / home server)

Run the OneNote MCP server as a container on any host you control — a VPS, a
home server, a NAS. You provide the public HTTPS endpoint (claude.ai requires
HTTPS); put the container behind a reverse proxy such as Caddy, nginx, or
Cloudflare Tunnel.

## Image

Released images are published to the GitHub Container Registry on every tag:

```
ghcr.io/ahmadalmezaal/onenote-mcp:latest
ghcr.io/ahmadalmezaal/onenote-mcp:0.2          # major.minor
ghcr.io/ahmadalmezaal/onenote-mcp:0.2.0        # exact version
```

Or build locally from the repo root: `docker build -t onenote-mcp .`

## 1. Sign in to OneNote locally

The container has no terminal for the device-code login, so run it once on
your laptop and carry the result over as a secret:

```sh
ONENOTE_MCP_CLIENT_ID=<your-client-id> npx @atomiclabs97/onenote-mcp login
```

## 2. Configure

Copy [`docker-compose.yml`](./docker-compose.yml) to your host and create a
`.env` file beside it:

```sh
ONENOTE_MCP_CLIENT_ID=<entra-app-client-id>
ONENOTE_MCP_HTTP_TOKEN=<paste output of: openssl rand -base64 32>
ONENOTE_MCP_TOKEN_CACHE=<paste verbatim contents of ~/.config/onenote-mcp/tokens.json>
# ONENOTE_MCP_TENANT_ID=         # optional, single-tenant Entra apps only
```

The server seeds its token cache from `ONENOTE_MCP_TOKEN_CACHE` on first boot,
then keeps the refreshed token on the `onenote-data` volume. Alternatively,
skip that variable and drop the file straight onto the volume at
`/data/onenote-mcp/tokens.json`.

## 3. Run

```sh
docker compose up -d
docker compose logs -f          # look for "HTTP server listening"
```

Verify locally:

```sh
curl http://localhost:3000/healthz                            # → 200
curl -X POST http://localhost:3000/mcp \
  -H "Authorization: Bearer $ONENOTE_MCP_HTTP_TOKEN" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

## 4. Expose over HTTPS

claude.ai will not connect to a plain-HTTP URL. Front the container with one of:

- **Caddy** — automatic Let's Encrypt; reverse-proxy `your.domain` → `localhost:3000`.
- **nginx / Traefik** — terminate TLS, proxy to `localhost:3000`.
- **Cloudflare Tunnel** — no open inbound ports; `cloudflared` maps a hostname to `localhost:3000`.

Keep `/healthz` reachable for upstream health checks; it needs no auth.

## 5. Connect claude.ai

Once `https://your.domain/mcp` is reachable, follow the
[claude.ai Connector walkthrough](./claude-ai-connector.md).

## Other platforms

Cloudflare Workers and other edge runtimes are **not supported**: the server
needs a Node runtime (`@azure/msal-node`, `node:http`) and a writable
filesystem for the token cache. Any host that can run the Docker image — or
Node 20+ directly — works.
