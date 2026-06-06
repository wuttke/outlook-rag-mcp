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
         ▲                                                              │
         │ live COM (attachments, drafts)                               ▼
         │                                                    ┌────────────────────┐
         │                                                    │  MCP server        │
         └────────────────────────────────────────────────────│  search_semantic   │
                                                              │  search_metadata   │
                                                              │  search_hybrid     │
                                                              │  get_full_email    │
                                                              │  list_attachments  │
                                                              │  read_attachment   │
                                                              │  create_draft      │
                                                              └────────────────────┘
```

- **Step 1** (`outlook_export.ps1`) — PowerShell + Outlook COM. Walks the
  default store, skips junk/calendar/contacts/etc., writes one mbox per folder,
  persists a `_sync_state.json` (per-folder watermark + seen-keys hash set) so
  re-runs are incremental. Also writes `_current_state.json`, a full snapshot
  of every mail's current folder + EntryID + unread/categories, so the next
  ingest can detect moves and deletes. Idempotent.
- **Step 2** (`ingest.py`) — chunks each mail body (2000 char windows, 200
  overlap), filters to human folders (configurable allowlist), embeds with
  `BAAI/bge-m3` on CPU, appends to LanceDB. Dedup key is
  `sha1(entry_id || chunk_idx)`, so re-runs only embed new chunks. Before
  appending, reconciles against the snapshot: drops rows whose mails no
  longer exist in any allowlisted folder, updates the folder column for
  mails that moved between allowlisted folders, and flips the per-row
  `unread` bit when the snapshot disagrees with the index.
- **Step 3** (`mcp_server.py`) — FastMCP server exposing seven tools.
  Vector/metadata search runs against the local LanceDB; attachment and
  draft tools talk to the live Outlook session via PowerShell + COM.
- **Driver** (`refresh.sh`) — runs step 1 in a "repeat until 0 new" loop (the
  COM `Restrict()` snapshot can miss items if OST sync is concurrent), then
  chains step 2.
- **Cron wrapper** (`cron-refresh.sh`) — runs the driver every 5 minutes
  under `flock`, skips overlapping runs, and caps the log file size.

## Layout

| File | Purpose |
|---|---|
| `outlook_export.ps1` | Step 1 — Outlook → mbox (Windows, via `powershell.exe`) |
| `ingest.py` | Step 2 — mbox → LanceDB |
| `mcp_server.py` | MCP tools (FastMCP) |
| `outlook_attachment.ps1` | Live-Outlook helper for attachment list/read |
| `outlook_draft.ps1` | Live-Outlook helper for draft creation / replies |
| `query.py` | Standalone CLI for ad-hoc search |
| `refresh.sh` | Driver: step 1 loop + step 2 |
| `cron-refresh.sh` | Cron-safe wrapper around `refresh.sh` (flock + log cap) |
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
| `search_semantic` | bge-m3 vector search + optional filters (incl. `unread`) |
| `search_metadata` | pure SQL filter (folder, sender, dates, AND-keywords, `unread`) |
| `search_hybrid` | Reciprocal Rank Fusion of semantic + keyword legs |
| `get_full_email` | reassemble all chunks of one mail by `entry_id` |
| `list_attachments` | live Outlook COM — list attachments on a mail (incl. inline) |
| `read_attachment` | live Outlook COM — fetch one attachment, base64-encoded |
| `create_draft` | live Outlook COM — compose a new mail or reply, saved as draft |

First call to any vector tool blocks ~25 s while bge-m3 loads; subsequent
calls are sub-second. The LanceDB table handle is refreshed on every MCP
request, so newly-ingested mails become searchable without restarting the
server. Attachment and draft tools require Outlook to be running on the
Windows side; `create_draft` only ever saves to the Drafts folder, it
never sends.

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
- The Outlook COM helpers drop the WSL process integrity level to Medium
  before binding, otherwise Office refuses the cross-integrity call.
- Snapshot reconciliation falls back to a no-op when `_current_state.json`
  is missing, so the pipeline never deletes rows spuriously.
- The `unread` column was added to existing tables via a one-time
  `add_columns` + `alter_columns` migration (no re-embedding); subsequent
  reconciliation passes keep it in sync with the live Outlook state.

## License

Licensed under the Apache License, Version 2.0. See [LICENSE](LICENSE).

Copyright 2026 Matthias Wuttke.
