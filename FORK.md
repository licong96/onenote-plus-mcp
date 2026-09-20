# FORK.md — onenote-plus-mcp

This is a **personal fork** of [`ahmadAlMezaal/onenote-mcp`](https://github.com/ahmadAlMezaal/onenote-mcp)
(MIT), branched from commit `894e82b`, and extended with a local page index and
seven tools that the upstream server does not have.

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
| `index` | Maintains the local mirror of the structure. Modes: `status`, `search` (offline), `sync` (incremental), `rebuild`, `rebuildSkeleton`. |
| `resolve_page` | Turns a page title (optionally with notebook/section) into a real page ID, served from the index and repaired against Graph on a miss. |
| `list_pages` | Lists pages in a section, or across every section of a notebook. Sorted server-side; no search term required. |
| `find_pages` | The working search. Served from the local index by default; falls back to a live scan only when the index is absent or does not know the requested scope. |
| `get_notebook_tree` | Maps notebooks → section groups → sections with IDs. The entry point for discovering every other ID. |
| `copy_page` | Copies a page into another section (`copyToSection`). Copy only — never moves or deletes. |
| `auth_status` | Which account is signed in, where the token cache lives, whether the cached token still refreshes. |

Plus supporting infrastructure:

- **Local page index** (`src/index-store/`) — mirrors notebooks, sections, and
  every page (ID, title, notebook, section, group path, timestamps, URL) to
  `~/.config/onenote-mcp/index.json`. This is the single biggest change; see the
  section below for why it is necessary.
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

## The local index, and why it exists

Measured on this account, **every Graph request costs 2–7 seconds regardless of
payload size** — listing a 7-page section took 7s, and listing a 10-section
notebook took 47s. The cost is the round trip, not the data. That makes anything
that fans out per section (172 sections here) both unusably slow and an instant
route to HTTP 429.

The structure is also almost entirely static — notebooks and sections rarely
change; only page *content* moves. So the skeleton and page lists are mirrored
to one JSON file and served from there:

```
~/.config/onenote-mcp/index.json   # 13 notebooks / 172 sections / 3087 pages ≈ 2.4 MB
```

Measured effect: searching "香港" across all 3087 pages went from **11181ms and a
429 error** to **19ms**.

### Two index keys, deliberately

- **byId** — exact, O(1), used once a page ID is known.
- **byPath** (notebook + section + title) — because **page IDs are not stable**.
  Moving or copying a page mints a new ID, so a stored ID can silently 404.
  The path key is the fallback that reconstructs it.

### Freshness

The index is a cache, not a source of truth. On a miss, resolution verifies
against Graph and **writes the repair back**, so the next lookup is cheap again.
Write paths keep it current automatically:

| Operation | Index action |
| --- | --- |
| `create_page` | re-sync that section |
| `delete_page` | drop the entry locally — no network call |
| `copy_page` | re-sync the target section (new page, new ID) |
| content edit (`update_page`) | no-op — metadata is unchanged |

All of these are **best-effort**: a failed refresh is reported as
`{status:'failed'}` but never fails the write itself, because the mutation has
already happened on Microsoft's side and rolling it back is not an option.

A full rebuild is ~185 requests and measured at **285s** at concurrency 3. Run
it once; afterwards prefer `sync`, scoped with `sections` when you know what
changed.

---

## Bugs found by actually running this

Both were invisible to mocked tests and only appeared against live Graph.

**1. Case-sensitivity mismatch in index filters.** The query needle was
lowercased but the `notebook`/`section` filters were not, so
`section: "Section B"` matched nothing while `section: "section b"` worked.
The bad outcome was silent — a normal-looking filter returned zero results and
read as "the page does not exist". Fixed by lowercasing both sides.

**2. `find_pages` resolved scope before checking the index.** Scope resolution
itself walks Graph (172 section requests with no hint), so the expensive work
ran *before* the cache was consulted — paying the exact cost the cache exists to
avoid, and tripping the rate limiter even on what should have been a cache hit.
Fixed by checking the index first and deriving scope from it via
`scopeFromIndexArgs`, with no network calls.

The lesson worth keeping: **a slow-path fallback must be checked *after* the
fast path, never the reverse.** A cache placed behind an expensive precondition
is not a cache.

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
npm test               # vitest, 299 tests (~96% statement coverage)
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
seven new files under `src/tools/`, one under `src/graph/`, one under `src/util/`,
four under `src/index-store/`, and small additive edits elsewhere (noted with
`--- fork additions ---` comments). `src/tools/index.ts`, `src/index.ts`,
`src/cli.ts`, and `package.json` carry the only renames.

---

*Created 2026-09-20.*

## Test coverage

The fork's additions are covered by `tests/util/concurrency.test.ts`,
`tests/graph/tree.test.ts`, `tests/graph/find.test.ts`,
`tests/graph/pagesList.test.ts`, `tests/tools/forkTools.test.ts`, and
`tests/index-store/` (store, sync, resolve, maintain), plus new cases appended to
`tests/config.test.ts` (config-dir override) and `tests/graph/client.test.ts`
(`paginateUntil`, `graphRequestRaw`).

Suite total: **299 tests**, 96% statements / 94% branches. The new modules are at
100% statement coverage apart from `copy_page`'s defensive branch and the
inherited `auth/index.ts` login/logout paths.
