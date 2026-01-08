# word add in dev

Deterministic, fast dev loop for Word/Excel add-ins on macOS.

This repo provides a single Node script that:
- validates the manifest (catches unescaped `&`)
- sideloads to the correct WEF folder
- optionally clears Office web caches (via `trash`)
- kills old servers, starts a new one, and health‑checks it
- optionally restarts Word/Excel

## Requirements

- macOS with Microsoft Word/Excel desktop
- Node.js 18+ (for built‑in `fetch`)
- `trash` command (optional, only if you use `--clear-cache true`)

## Quick start

From your add‑in project directory:

```bash
node /path/to/word-addin-dev/word-addin-dev.js \
  --manifest ./manifest.xml \
  --server-cmd "node dist/your-app.js serve --port 4321" \
  --app word \
  --port 4321
```

This will:
1) validate the manifest  
2) sideload it  
3) kill anything on port 4321  
4) start your server  
5) wait for `http://localhost:4321/api/health`  
6) restart Word

## Common variants

Code-only change (no manifest, no restart):

```bash
node /path/to/word-addin-dev/word-addin-dev.js \
  --server-cmd "node dist/your-app.js serve --port 4321" \
  --sideload false \
  --restart false
```

Full deterministic reset (includes cache clear):

```bash
node /path/to/word-addin-dev/word-addin-dev.js \
  --manifest ./manifest.xml \
  --server-cmd "node dist/your-app.js serve --port 4321" \
  --clear-cache true
```

## Options

- `--manifest <path>`: XML manifest to validate + sideload
- `--server-cmd <command>`: server command (quoted)
- `--health-url <url>`: health check (default: `http://localhost:4321/api/health`)
- `--app word|excel|both`: target app (default: `word`)
- `--port <n>`: port used for kill‑port (default: `4321`)
- `--pid-file <path>`: pid file for server (default: `.cache/verify/server.pid`)
- `--log-file <path>`: server log (default: `.cache/verify/server.log`)
- `--sideload true|false`: copy manifest into `wef` (default: true if manifest provided)
- `--restart true|false`: restart app (default: true if sideload true)
- `--clear-cache true|false`: clear Office web cache via `trash`
- `--validate true|false`: validate manifest (default: true)
- `--health true|false`: wait for health (default: true)
- `--kill-port true|false`: kill listeners on port (default: true)

## What this does not do

- It does **not** generate a manifest.
- It does **not** bundle/build your add‑in.
- It does **not** change your taskpane URL.

## Deterministic workflow reference

- HTML/JS changes only → Clear Web Cache → Reload (no restart)
- Manifest changes → quit + reopen Word
- Sideload changes → copy manifest to `wef` + restart Word

This script automates the deterministic path.

