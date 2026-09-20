# Deploy to Fly.io

Get a public, HTTPS OneNote MCP endpoint that [claude.ai](https://claude.ai)
(web + iOS/Android) can connect to. Free-tier eligible; ~10 minutes.

## Prerequisites

- A Microsoft Entra app registration — see the [README Quick start](../../README.md#1-register-a-microsoft-entra-app) for the 2-minute walkthrough. You need its **Application (client) ID**.
- [`flyctl`](https://fly.io/docs/flyctl/install/) installed and `fly auth login` done.
- A local checkout of this repo (the deploy builds the bundled `Dockerfile`).

## 1. Sign in to OneNote locally

The hosted server has no terminal, so the device-code login runs **once on your
laptop**. Its output — `tokens.json` — is uploaded to Fly as a secret.

```sh
ONENOTE_MCP_CLIENT_ID=<your-client-id> npx @atomiclabs97/onenote-mcp login
```

This caches a refresh token at `~/.config/onenote-mcp/tokens.json`.

## 2. Launch the app

From the repo root:

```sh
fly launch --no-deploy
```

Pick a unique app name when prompted.

Create the volume the config expects:

```sh
fly volumes create onenote_data --size 1 --region iad
```

## 3. Set the secrets

```sh
fly secrets set \
  ONENOTE_MCP_CLIENT_ID="<your-client-id>" \
  ONENOTE_MCP_HTTP_TOKEN="$(openssl rand -base64 32)" \
  ONENOTE_MCP_TOKEN_CACHE="$(cat ~/.config/onenote-mcp/tokens.json)"
```

- `ONENOTE_MCP_HTTP_TOKEN` — the bearer token claude.ai must send. **Copy the
  generated value** (`fly secrets list` only shows digests). You need it in
  step 5.
- `ONENOTE_MCP_TOKEN_CACHE` — the verbatim `tokens.json`. On first boot the
  server seeds its cache from this, then refreshes the token on the `/data`
  volume from then on. (Optional: add `ONENOTE_MCP_TENANT_ID` for a
  single-tenant Entra app.)

## 4. Deploy

```sh
fly deploy
```

When it finishes, your endpoint is `https://<app-name>.fly.dev`. Verify:

```sh
curl https://<app-name>.fly.dev/healthz                       # → 200, empty body

curl -X POST https://<app-name>.fly.dev/mcp \
  -H "Authorization: Bearer <your-http-token>" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
```

The second call should return the 11 OneNote tools. A `401` means the bearer
token is wrong; a Graph auth error means the token cache didn't seed — check
`fly logs` for `Seeded OneNote token cache`.

## 5. Connect claude.ai

Continue with the [claude.ai Connector walkthrough](./claude-ai-connector.md)
using the `https://<app-name>.fly.dev/mcp` URL and the bearer token from step 3.

## Rotating the bearer token

```sh
fly secrets set ONENOTE_MCP_HTTP_TOKEN="$(openssl rand -base64 32)"
```

Fly redeploys automatically. Update the token in the claude.ai Connector too.

## Re-authenticating OneNote

If the refresh token is revoked or expires, repeat step 1 locally, then
re-set `ONENOTE_MCP_TOKEN_CACHE` and clear the stale volume copy:

```sh
fly ssh console -C "rm -f /data/onenote-mcp/tokens.json"
fly secrets set ONENOTE_MCP_TOKEN_CACHE="$(cat ~/.config/onenote-mcp/tokens.json)"
```
