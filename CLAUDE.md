# CLAUDE.md

Project conventions for Claude Code (and any other AI assistant working in this repo). Read this before making changes.

## What this project is

`@atomiclabs97/onenote-mcp` — an open-source MCP (Model Context Protocol) server for **Microsoft OneNote**. Users run it via `npx @atomiclabs97/onenote-mcp` and wire it into Claude Desktop, Cursor, or any MCP-compatible client. Auth uses MSAL device-code against a user-owned Microsoft Entra app registration. The server exposes typed tools that talk to Microsoft Graph.

## Tech stack

- **TypeScript** with strict mode, ES2022 target, NodeNext module resolution
- **Node 20+** (required for stable global `fetch`)
- **Yarn (Classic, 1.x)** — _not_ npm. Lockfile is `yarn.lock`. Do not commit `package-lock.json`.
- `@modelcontextprotocol/sdk` for the MCP server
- `@azure/msal-node` for OAuth
- `zod` (`^4.x`) for tool input schemas
- `marked` + `turndown` for MD ↔ HTML
- `vitest` for tests, `eslint` (flat config) for linting

## Hard conventions

| Rule | Why | Enforced by |
| --- | --- | --- |
| Arrow functions for all module-scope functions (no `function foo() {}`) | Project style | ESLint `func-style: ['error', 'expression']` |
| No `package-lock.json` in repo | Yarn is the source of truth | Convention (don't run `npm install`) |
| Conventional commits (`feat:`, `fix:`, `chore:`, `docs:`, `refactor:`, `test:`, `ci:`) | Readable history, release tooling | Convention |
| Each tool in its own file under `src/tools/` exporting `register(server)` | Discoverability | Convention |
| Tools never call `fetch` directly — they go through `src/graph/*` | Single retry / auth / error-shaping path | Convention |
| Cross-directory imports use the `@/*` alias (`@/graph/pages.js`), not `../../graph/pages.js`. Sibling imports stay `./foo.js`. | Readability, refactor-safe | `tsconfig.json` paths + `tsc-alias` at build, `vitest.config.ts` for tests |
| Default to writing **no comments**. Only the WHY when non-obvious. | Self-documenting code | Convention |
| Don't add error handling, fallbacks, or validation for cases that can't happen | Trust internal contracts | Convention |

## Project structure

```
src/
  index.ts                # MCP server factory + stdio runner
  cli.ts                  # CLI dispatch: login | logout | serve (default)
  config.ts               # env vars, XDG paths, scopes, authority
  markdown.ts             # md ↔ html helpers (full doc + fragment)
  auth/                   # MSAL client + persistent token cache
  graph/                  # Typed Microsoft Graph wrappers (one file per resource)
  http/                   # Streamable HTTP transport (server, auth gate, healthz)
  tools/                  # One file per MCP tool + index.ts that registers them
tests/                    # Vitest suites mirroring src/
scripts/                  # Maintainer scripts (e.g. smoke.ts) — not shipped to npm
.github/workflows/        # CI (typecheck/lint/test/build) + release (npm publish + Docker image on tag)
docs/deployment/          # Self-hosting guides (Fly.io, Docker, claude.ai Connector)
docs/images/              # Entra walkthrough screenshots referenced by README
```

## Dev loop

```sh
yarn install
yarn typecheck   # tsc --noEmit, covers src/ + scripts/
yarn lint        # eslint .
yarn test        # vitest run (unit tests, mocked Graph)
yarn build       # tsc + tsc-alias → dist/
yarn dev         # incremental rebuild: tsc --watch + tsc-alias --watch
yarn smoke       # end-to-end against a real OneNote account (needs login first)
```

To test the CLI locally against Claude Desktop:

1. `yarn build`
2. Point Claude Desktop's `mcpServers.onenote.command` at `node`, `args` at the absolute path of `dist/cli.js`, and set `env.ONENOTE_MCP_CLIENT_ID`.
3. Restart Claude Desktop.

## Auth setup

The end-to-end auth flow is documented in [`README.md`](README.md#1-register-a-microsoft-entra-app) — Entra app registration, device-code login, token caching. Don't paste that walkthrough into source comments; link to the README.

## Adding a new tool

1. Add a typed Graph wrapper in `src/graph/*.ts`. No `fetch` calls outside this directory.
2. Add `src/tools/<toolName>.ts` exporting `const register = (server: McpServer): void => { server.registerTool(...) }`. Import the Graph wrapper as `from '@/graph/<resource>.js'`.
3. Wire it into `src/tools/index.ts`.
4. Add a row to the tool reference table in `README.md`.
5. Add tests for any non-trivial logic (parsing, format conversion, validation).
6. Run `yarn typecheck && yarn lint && yarn test` before opening the PR.

## Don't

- Don't introduce new dependencies without flagging it in the PR description.
- Don't write multi-paragraph code comments or docstrings. One short line max.
- Don't commit `package-lock.json`. Don't switch the project back to npm.
- Don't write `function foo() {}` at module scope. Use `const foo = () => {}`.
- Don't reach for the bare `onenote-mcp` npm name — it's taken. We publish as `@atomiclabs97/onenote-mcp`.
- Don't propose roadmap features that are already tracked. Check the **Roadmap** section in `README.md` first.
- Don't reformat or restyle code unrelated to the change you're making.
- Don't use `package-lock.json` syntax or `npm run <script>` in docs. Always `yarn <script>`.

## Where to look next

- **Tool reference** — [`README.md` § Tools](README.md#tools)
- **Roadmap and known limitations** — [`README.md` § Roadmap](README.md#roadmap) and [§ Known limitations](README.md#known-limitations)
- **Contribution policy** — [`CONTRIBUTING.md`](CONTRIBUTING.md)
