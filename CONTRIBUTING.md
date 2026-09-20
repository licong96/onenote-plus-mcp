# Contributing to onenote-mcp

Thanks for considering a contribution. The project is intentionally small — please read the [README](README.md) first to understand the v1 scope.

## Development loop

```sh
yarn install
yarn typecheck   # tsc --noEmit
yarn lint        # eslint
yarn test        # vitest
yarn build       # tsc + tsc-alias → dist/
```

> Yarn (Classic, 1.x) is the source of truth for the lockfile — do not commit `package-lock.json`.

Run the built CLI locally:

```sh
ONENOTE_MCP_CLIENT_ID=<your-client-id> node dist/cli.js login
ONENOTE_MCP_CLIENT_ID=<your-client-id> node dist/cli.js          # serve
```

To wire your local checkout into Claude Desktop or Cursor while iterating, point the MCP client `command` at `node` and `args` at the absolute path of `dist/cli.js`.

## Commit style

We use [Conventional Commits](https://www.conventionalcommits.org/):

- `feat: …` — user-visible new functionality
- `fix: …` — bug fix
- `chore: …` — tooling, deps, scaffolding
- `docs: …` — README / docs only
- `refactor: …` — internal change, no behavior delta
- `test: …` — test additions / changes
- `ci: …` — workflow changes

Keep the subject under 72 characters. Wrap the body at ~80.

## Pull requests

1. Open an issue first for anything non-trivial (new tool, behavior change, dep addition).
2. Each PR should be focused and pass `typecheck`, `lint`, and `test` locally.
3. Include a short rationale in the PR description — what changes, why, and any tradeoffs.

## Adding a new tool

1. Add a typed Graph wrapper in `src/graph/`.
2. Add a tool file in `src/tools/` exporting `register(server)`. Import the wrapper via the `@/*` alias (e.g. `import { foo } from '@/graph/foo.js'`) — cross-directory imports use the alias, sibling imports stay `./foo.js`.
3. Wire it into `src/tools/index.ts`.
4. Update the tool reference table in `README.md`.
5. Add tests for any non-trivial logic (input validation, format handling, etc.).

## Releasing

Releases are cut by pushing a tag. The `release.yml` workflow publishes to npm using the `NPM_TOKEN` repo secret and pushes a Docker image to the GitHub Container Registry.

```sh
yarn version --patch   # or --minor / --major
git push --follow-tags
```
