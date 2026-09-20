# FORK.md — onenote-plus-mcp

This is a **personal fork** of [`ahmadAlMezaal/onenote-mcp`](https://github.com/ahmadAlMezaal/onenote-mcp)
(MIT), branched from commit `894e82b`, and extended with five tools that the
upstream server does not have.

The upstream project stays upstream. Everything below is what this fork adds and
*why*, so the reasoning survives past the session that produced it.

---

## Why the fork exists

Two concrete gaps, both confirmed against the live account (13 notebooks,
172 sections) rather than guessed:

1. **No way to list pages.** There is no `list_pages` upstream. `search_pages` is
   the only way to discover a page, and it needs a keyword you already know.
2. **Search is broken on this account.** Graph rejects `$search` here:

   ```text
   GET /me/onenote/pages?$search="香港"      → Your request contains unsupported OData query parameters.
   GET /me/onenote/pages                     → The number of maximum sections is exceeded for this request.
   GET /me/onenote/sections/{id}/pages?$search="a"
                                             → Your request contains unsupported OData query parameters.
   ```

   The account-wide `/me/onenote/pages` collection is unusable at 172 sections,
   and `$search` is rejected whether or not the query is section-scoped. So the
   upstream `search_pages` tool cannot work here, and neither can anything else
   built on those two endpoints.

   (`search_pages` is therefore excluded from the OpenClaw tool filter. The
   upstream docs claimed search worked; that is no longer true for this account.)

---

## What was added

| Tool | Purpose |
| --- | --- |
| `list_pages` | Lists pages in a section, or across every section of a notebook. Sorted server-side; no search term required. |
| `find_pages` | The working search. Scans sections in scope and matches client-side — titles by default, page bodies on request. |
| `get_notebook_tree` | Maps notebooks → section groups → sections with IDs. The entry point for discovering every other ID. |
| `copy_page` | Copies a page into another section (`copyToSection`). Copy only — never moves or deletes. |
| `auth_status` | Which account is signed in, where the token cache lives, whether the cached token still refreshes. |

Plus supporting infrastructure:

- `paginateUntil` (`src/graph/client.ts`) — paginate but stop once `limit` rows
  are collected, so "10 newest pages of a 500-page section" is one request.
- `graphRequestRaw` (`src/graph/client.ts`) — same auth/retry/error shaping as
  `graphRequest`, but returns status + headers + body. Needed because
  `copyToSection` answers `202 Accepted` with the job URL in a header.
- `mapWithConcurrency` (`src/util/concurrency.ts`) — bounded fan-out. A
  whole-account scan is 170+ requests: sequential is too slow, unbounded trips
  the throttle.
- `ONENOTE_MCP_CONFIG_DIR` — relocates the token cache, so this fork and the
  upstream server can hold independent logins when run side by side.

---

## Design decisions worth keeping

- **No `$orderby` on section / section-group collections.** The OneNote API
  accepts it on some endpoints and rejects the whole request with "unsupported
  OData query parameters" on others. Those lists are small — they are sorted
  client-side instead.
- **Notebook-level page listings fan out per section.** Graph has no
  notebook→pages endpoint. Each section is asked for its own top *N*; any page
  belonging in the notebook-wide top *N* must be in its section's top *N*, so
  the merge is exact without downloading whole sections.
- **`collectSectionsForNotebook` walks section groups.** `/notebooks/{id}/sections`
  alone misses sections nested inside section groups, which would silently
  truncate every notebook-wide scan. (This account happens to have no section
  groups, so that path is not exercised by live data — it is written from the
  documented model, not from observation.)
- **Content search is opt-in and capped.** Titles come free with page metadata;
  bodies cost one download each, so `includeContent` defaults to `false` and the
  scan stops at `maxContentPages` (default 100) and says so in `notes`.
- **`copy_page` resolves both ends before issuing the copy.** A bad ID fails
  before an async job is queued, and the response names real pages instead of
  raw GUIDs.
- **`auth_status` never throws.** A diagnostic that fails is useless exactly
  when you need it; a broken cache reports `signedIn: false`.

---

## Layout

```text
/Users/liangluyang/Desktop/learning/ai/
├── onenote-mcp-upstream/   # pristine clone of the original repo (reference)
└── onenote-plus-mcp/       # this fork — build and run from here
```

## Build and run

```bash
cd /Users/liangluyang/Desktop/learning/ai/onenote-plus-mcp
npm install
npm run build          # tsc + tsc-alias  → dist/
npm run typecheck
npm test               # vitest, 245 tests (~96% statement coverage)
node dist/cli.js       # stdio MCP server
```

`yarn` is not installed on this machine; the upstream `prepublishOnly` hook was
switched to `npm run build` accordingly.

Login reuses the upstream token cache at `~/.config/onenote-mcp/tokens.json`,
so no second sign-in is needed. To keep them separate:

```bash
ONENOTE_MCP_CLIENT_ID=<id> ONENOTE_MCP_CONFIG_DIR=onenote-plus-mcp \
  node dist/cli.js login
```

## How OpenClaw uses it

`onenoteplus` in `~/.openclaw/openclaw.json`, launched as

```text
node /Users/liangluyang/Desktop/learning/ai/onenote-plus-mcp/dist/cli.js
```

with `ONENOTE_MCP_CLIENT_ID=14d82eec-204b-4c2f-b7e8-296a70dab67e` (Microsoft
Graph Command Line Tools — no Entra app registration required) and
`toolFilter.exclude: ["search_pages"]`.

The old `onenote` entry is kept but disabled (`enabled: false`) rather than
deleted. **To roll back:**

```bash
openclaw mcp configure onenote --enable      # restore the upstream server
openclaw mcp configure onenoteplus --disable # stand this one down
```

## Syncing with upstream

`onenote-mcp-upstream/` is the untouched clone. New work here is additive:
five new files under `src/tools/`, one under `src/graph/`, one under `src/util/`,
and small additive edits elsewhere (noted with `--- fork additions ---`
comments). `src/tools/index.ts`, `src/index.ts`, `src/cli.ts`, and `package.json`
carry the only renames.

---

*Created 2026-09-20.*

## Test coverage

The fork's additions are covered by `tests/util/concurrency.test.ts`,
`tests/graph/tree.test.ts`, `tests/graph/find.test.ts`,
`tests/graph/pagesList.test.ts`, and `tests/tools/forkTools.test.ts`, plus new
cases appended to `tests/config.test.ts` (config-dir override) and
`tests/graph/client.test.ts` (`paginateUntil`, `graphRequestRaw`).

Suite total: 245 tests, 96% statements / 94% branches. The new modules are at
100% statement coverage apart from `copy_page`'s defensive branch and the
inherited `auth/index.ts` login/logout paths.
