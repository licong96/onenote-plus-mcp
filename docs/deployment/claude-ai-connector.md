# Connect to claude.ai

Once your OneNote MCP server is deployed and reachable over HTTPS (see
[Fly.io](./fly.md) or [Docker](./docker.md)), add it to claude.ai as a custom
Connector. Works on web and the iOS/Android apps.

## What you need

- The MCP endpoint URL — `https://<your-host>/mcp`.
- The bearer token — the value you set as `ONENOTE_MCP_HTTP_TOKEN`.

## Add the Connector

1. Open [claude.ai](https://claude.ai) → **Settings** → **Connectors**.
2. Click **Add custom connector**.
3. Fill in:
   - **Name** — e.g. `OneNote`.
   - **Remote MCP server URL** — `https://<your-host>/mcp`.
   - **Authentication** — choose the bearer-token / custom-header option and set
     the header to `Authorization: Bearer <your-http-token>`.
4. Save. claude.ai connects and lists the server's tools.

## Verify

In a new chat, open the tool/connector picker and confirm the 11 OneNote tools
appear: `list_notebooks`, `list_sections`, `list_section_groups`,
`search_pages`, `read_page`, `create_notebook`, `create_section`,
`create_section_group`, `create_page`, `update_page`, `delete_page`.

End-to-end smoke test — ask Claude:

> Create a page titled "Connector test" in my OneNote with a short note, then
> read it back.

It should call `create_page` then `read_page` and return the content. Try it
from the mobile app too — the same Connector is shared across devices.

## Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| Connector fails to add / "could not reach server" | URL not HTTPS, wrong path (must end in `/mcp`), or server down — check `/healthz`. |
| Tools listed but every call errors with auth | Server can't reach Graph — the token cache didn't seed or the refresh token expired. Re-do the login step in the deployment guide. |
| `401` on connect | Bearer token mismatch between the Connector and `ONENOTE_MCP_HTTP_TOKEN`. |

## Notes

- One deployment serves **one OneNote account** — whichever account the seeded
  `tokens.json` was created for. Run a separate instance per account.
- Anyone with the URL **and** the bearer token can read and write that OneNote.
  Treat the token like a password; rotate it if leaked (see the deployment guide).

> Screenshots of the claude.ai Connector screens are intentionally omitted —
> the settings UI changes often enough that step text stays accurate longer.
