# outlook-rag-mcp

Local RAG over an Outlook mailbox: export from the Windows OST to mbox, embed
into LanceDB with `BAAI/bge-m3`, then query via an MCP server with semantic,
metadata, and hybrid (RRF) search.

Everything runs locally. No mail or embeddings leave the host.

## Architecture

```
┌──────────────────┐  step 1  ┌───────────────────────┐  step 2  ┌──────────────┐
│  Outlook (OST)   │ ───────▶ │  *.mbox (per folder)  │ ───────▶ │   LanceDB    │
│  Windows / MAPI  │   PS1    │  C:\…\outlook-export  │  Python  │  ~/db (1024d)│
└──────────────────┘          └───────────────────────┘          └──────┬───────┘
                                                                        │
                                                                        ▼
                                                              ┌──────────────────┐
                                                              │  MCP server      │
                                                              │  search_semantic │
                                                              │  search_metadata │
                                                              │  search_hybrid   │
                                                              │  get_full_email  │
                                                              └──────────────────┘
```

- **Step 1** (`outlook_export.ps1`) — PowerShell + Outlook COM. Walks the
  default store, skips junk/calendar/contacts/etc., writes one mbox per folder,
  persists a `_sync_state.json` (per-folder watermark + seen-keys hash set) so
  re-runs are incremental. Idempotent.
- **Step 2** (`ingest.py`) — chunks each mail body (2000 char windows, 200
  overlap), filters to human folders (configurable allowlist), embeds with
  `BAAI/bge-m3` on CPU, appends to LanceDB. Dedup key is
  `sha1(entry_id || chunk_idx)`, so re-runs only embed new chunks.
- **Step 3** (`mcp_server.py`) — FastMCP server exposing four tools.
- **Driver** (`refresh.sh`) — runs step 1 in a "repeat until 0 new" loop (the
  COM `Restrict()` snapshot can miss items if OST sync is concurrent), then
  chains step 2.

## Layout

| File | Purpose |
|---|---|
| `outlook_export.ps1` | Step 1 — Outlook → mbox (Windows, via `powershell.exe`) |
| `ingest.py` | Step 2 — mbox → LanceDB |
| `mcp_server.py` | MCP tools (FastMCP) |
| `query.py` | Standalone CLI for ad-hoc search |
| `refresh.sh` | Driver: step 1 loop + step 2 |
| `db/` *(not in git)* | LanceDB tables |
| `logs/` *(not in git)* | Per-run export and ingest logs |

## Setup (WSL / Linux)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install sentence-transformers lancedb pyarrow tqdm fastmcp
```

bge-m3 weights (~2.3 GB) are downloaded on first run.

## Usage

End-to-end refresh:

```bash
./refresh.sh                  # step 1 (loop to convergence) + step 2
./refresh.sh --no-ingest      # just refresh mboxes
./refresh.sh --max-passes 3   # cap the export loop
```

Manual ingest:

```bash
source .venv/bin/activate
python3 ingest.py --batch 16                       # default human folders
python3 ingest.py --folders Archiv,Posteingang     # custom allowlist
python3 ingest.py --reset                          # drop & rebuild table
```

Manual export (Windows side, from WSL):

```bash
powershell.exe -ExecutionPolicy Bypass -File "$(wslpath -w ./outlook_export.ps1)"
```

MCP server:

```bash
source .venv/bin/activate
python3 mcp_server.py
```

Register in `~/.augment/settings.json` (or any other MCP client) as a stdio
server pointing at `mcp_server.py`.

## Tools exposed by the MCP server

| Tool | Use it for |
|---|---|
| `search_semantic` | bge-m3 vector search + optional filters |
| `search_metadata` | pure SQL filter (folder, sender, dates, AND-keywords) |
| `search_hybrid` | Reciprocal Rank Fusion of semantic + keyword legs |
| `get_full_email` | reassemble all chunks of one mail by `entry_id` |

First call to any vector tool blocks ~25 s while bge-m3 loads; subsequent
calls are sub-second.

## Configuration

Paths and the folder allowlist are constants at the top of each script:

- `ingest.py`: `EXPORT_DIR`, `DB_PATH`, `HUMAN_FOLDERS`, chunk sizes.
- `outlook_export.ps1`: `$OutputDir`, `$SkipRoles`, `$SkipNames`.

The export script writes **every** mail folder (including tool noise like
Bitbucket/Nagios/JIRA). `ingest.py` filters by `X-Outlook-Folder` exact match
against `HUMAN_FOLDERS`, so subfolders are not implicitly included.

## Notes

- Single-user, single-host design. No multi-tenant story, no auth.
- Embeddings live in `db/` and are git-ignored. The model is multilingual
  (good for mixed German/English archives).
- Re-running export or ingest is always safe; both are idempotent.
- The exporter is best-effort: OST sync races and COM watermark quirks mean a
  single pass can miss items. `refresh.sh` works around this by looping.

## License

Licensed under the Apache License, Version 2.0. See [LICENSE](LICENSE).

Copyright 2026 Matthias Wuttke.
