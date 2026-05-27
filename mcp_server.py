"""
MCP server exposing the Outlook RAG store.

Tools:
  - search_semantic: bge-m3 vector search (German + English), optional metadata filters
  - search_metadata: SQL-style filter search (folder, sender, date range, keywords)
  - search_hybrid:   RRF fusion of semantic + keyword/metadata results
  - get_full_email:  fetch all chunks for a single email by entry_id, reconstructed

Run (stdio transport, for Claude Desktop / VSCode-style MCP clients):
    python mcp_server.py
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any

import lancedb
from mcp.server.fastmcp import FastMCP
from sentence_transformers import SentenceTransformer

DB_DIR = Path(os.environ.get("OUTLOOK_RAG_DB", str(Path.home() / "outlook-rag" / "db")))
TABLE_NAME = "messages"
MODEL_NAME = "BAAI/bge-m3"

mcp = FastMCP("outlook-rag")

_model: SentenceTransformer | None = None
_table = None

PROJECT_COLS = [
    "entry_id", "chunk_idx", "folder", "subject",
    "from_addr", "to_addr", "cc_addr", "date",
    "message_id", "attachments", "body_chunk",
]


def model() -> SentenceTransformer:
    global _model
    if _model is None:
        _model = SentenceTransformer(MODEL_NAME, device="cpu")
        _model.max_seq_length = 512
    return _model


def table():
    global _table
    if _table is None:
        db = lancedb.connect(str(DB_DIR))
        _table = db.open_table(TABLE_NAME)
    else:
        # Pick up rows appended by ingest.py since the table was first opened.
        _table.checkout_latest()
    return _table


def _esc(s: str) -> str:
    return s.replace("'", "''")


def _build_where(
    folder: str | None = None,
    from_addr: str | None = None,
    to_addr: str | None = None,
    subject_contains: str | None = None,
    since: str | None = None,
    until: str | None = None,
    keywords: list[str] | None = None,
) -> str | None:
    f: list[str] = []
    if folder:
        f.append(f"folder = '{_esc(folder)}'")
    if from_addr:
        f.append(f"from_addr LIKE '%{_esc(from_addr)}%'")
    if to_addr:
        f.append(f"to_addr LIKE '%{_esc(to_addr)}%'")
    if subject_contains:
        f.append(f"subject LIKE '%{_esc(subject_contains)}%'")
    if since:
        f.append(f"date >= timestamp '{since}T00:00:00Z'")
    if until:
        f.append(f"date <= timestamp '{until}T23:59:59Z'")
    if keywords:
        for kw in keywords:
            e = _esc(kw)
            f.append(f"(subject LIKE '%{e}%' OR body_chunk LIKE '%{e}%')")
    return " AND ".join(f) if f else None


def _format(r: dict, score: float | None = None) -> dict[str, Any]:
    body = (r.get("body_chunk") or "").strip()
    if len(body) > 800:
        body = body[:800] + " …"
    out = {
        "folder": r.get("folder"),
        "date": r["date"].isoformat() if r.get("date") else None,
        "from": r.get("from_addr"),
        "to": r.get("to_addr"),
        "subject": r.get("subject"),
        "chunk_idx": r.get("chunk_idx"),
        "entry_id": r.get("entry_id"),
        "body": body,
    }
    if score is not None:
        out["score"] = round(score, 4)
    return out


def _dedup_by_email(rows: list[dict], k: int) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for r in rows:
        eid = r.get("entry_id") or ""
        if eid in seen:
            continue
        seen.add(eid)
        out.append(r)
        if len(out) >= k:
            break
    return out


@mcp.tool()
def search_semantic(
    query: str,
    k: int = 8,
    folder: str | None = None,
    from_addr: str | None = None,
    to_addr: str | None = None,
    since: str | None = None,
    until: str | None = None,
) -> list[dict[str, Any]]:
    """Semantic search over Outlook mail using bge-m3 (multilingual).

    Args:
        query: natural-language question or topic
        k: number of results to return
        folder: exact folder name (e.g. Archiv, Posteingang, Gesendete Elemente)
        from_addr: substring matched against sender email
        to_addr: substring matched against recipient email
        since: ISO date YYYY-MM-DD (inclusive lower bound on Date header)
        until: ISO date YYYY-MM-DD (inclusive upper bound)

    Returns a list of hits (most similar first) with score in [0,1].
    """
    qvec = model().encode([query], normalize_embeddings=True)[0].tolist()
    builder = table().search(qvec).metric("cosine")
    where = _build_where(folder=folder, from_addr=from_addr, to_addr=to_addr,
                         since=since, until=until)
    if where:
        builder = builder.where(where, prefilter=True)
    rows = builder.limit(k * 3).to_list()
    deduped = _dedup_by_email(rows, k)
    return [_format(r, score=1.0 - r.get("_distance", 0.0)) for r in deduped]


@mcp.tool()
def search_metadata(
    keywords: list[str] | None = None,
    folder: str | None = None,
    from_addr: str | None = None,
    to_addr: str | None = None,
    subject_contains: str | None = None,
    since: str | None = None,
    until: str | None = None,
    k: int = 20,
    sort: str = "date_desc",
) -> list[dict[str, Any]]:
    """Filter-based search (no embeddings).

    All filters are AND-combined. Each keyword must appear (case-sensitive
    substring) in subject OR body chunk.

    Args:
        keywords: list of strings; each must appear in subject or body
        folder: exact folder name
        from_addr: substring matched against sender
        to_addr: substring matched against recipient
        subject_contains: substring matched against subject
        since: ISO date YYYY-MM-DD
        until: ISO date YYYY-MM-DD
        k: number of distinct emails to return
        sort: 'date_desc' (default) or 'date_asc'
    """
    where = _build_where(folder=folder, from_addr=from_addr, to_addr=to_addr,
                         subject_contains=subject_contains, since=since,
                         until=until, keywords=keywords)
    ds = table().to_lance()
    scanner = ds.scanner(
        filter=where,
        columns=PROJECT_COLS,
        limit=max(k * 10, 200),
    )
    rows = scanner.to_table().to_pylist()
    reverse = sort != "date_asc"
    rows.sort(key=lambda r: r.get("date") or datetime.min, reverse=reverse)
    deduped = _dedup_by_email(rows, k)
    return [_format(r) for r in deduped]


@mcp.tool()
def get_full_email(entry_id: str) -> dict[str, Any]:
    """Return a single email reassembled from all its chunks.

    Args:
        entry_id: Outlook X-Outlook-EntryID (returned by other tools as `entry_id`)

    Returns metadata plus the full body text (chunks joined in order).
    """
    where = f"entry_id = '{_esc(entry_id)}'"
    ds = table().to_lance()
    rows = ds.scanner(filter=where, columns=PROJECT_COLS).to_table().to_pylist()
    if not rows:
        return {"error": f"no email found for entry_id={entry_id!r}"}
    rows.sort(key=lambda r: r.get("chunk_idx") or 0)
    head = rows[0]
    full_body = "\n".join((r.get("body_chunk") or "").strip() for r in rows).strip()
    return {
        "entry_id": head.get("entry_id"),
        "folder": head.get("folder"),
        "date": head["date"].isoformat() if head.get("date") else None,
        "from": head.get("from_addr"),
        "to": head.get("to_addr"),
        "cc": head.get("cc_addr"),
        "subject": head.get("subject"),
        "message_id": head.get("message_id"),
        "attachments": head.get("attachments"),
        "num_chunks": len(rows),
        "body": full_body,
    }


@mcp.tool()
def search_hybrid(
    query: str,
    keywords: list[str] | None = None,
    k: int = 8,
    folder: str | None = None,
    from_addr: str | None = None,
    to_addr: str | None = None,
    since: str | None = None,
    until: str | None = None,
    rrf_k: int = 60,
) -> list[dict[str, Any]]:
    """Reciprocal-Rank-Fusion of semantic and keyword/metadata search.

    Runs two retrievals over the same filtered subset and fuses ranks:
      semantic: bge-m3 vector similarity to `query`
      keyword:  metadata scan, ordered by date desc, filtered by `keywords`
                (each keyword must appear in subject or body)
    Fusion score = sum_per_list( 1 / (rrf_k + rank) ).

    Args:
        query: natural-language query for the semantic leg
        keywords: AND-matched substrings for the keyword leg (defaults to [query] if None)
        k: number of distinct emails to return
        folder, from_addr, to_addr, since, until: filters applied to both legs
        rrf_k: RRF dampening constant (default 60)
    """
    if not keywords:
        keywords = [query]
    pool = max(k * 5, 30)

    # --- semantic leg ---
    qvec = model().encode([query], normalize_embeddings=True)[0].tolist()
    sb = table().search(qvec).metric("cosine")
    where_sem = _build_where(folder=folder, from_addr=from_addr, to_addr=to_addr,
                             since=since, until=until)
    if where_sem:
        sb = sb.where(where_sem, prefilter=True)
    sem_rows = _dedup_by_email(sb.limit(pool * 3).to_list(), pool)

    # --- keyword leg ---
    where_kw = _build_where(folder=folder, from_addr=from_addr, to_addr=to_addr,
                            since=since, until=until, keywords=keywords)
    ds = table().to_lance()
    kw_rows = ds.scanner(filter=where_kw, columns=PROJECT_COLS,
                         limit=pool * 10).to_table().to_pylist()
    kw_rows.sort(key=lambda r: r.get("date") or datetime.min, reverse=True)
    kw_rows = _dedup_by_email(kw_rows, pool)

    # --- RRF fusion by entry_id ---
    fused: dict[str, dict[str, Any]] = {}
    for rank, r in enumerate(sem_rows, 1):
        eid = r.get("entry_id") or ""
        fused.setdefault(eid, {"row": r, "score": 0.0, "src": []})
        fused[eid]["score"] += 1.0 / (rrf_k + rank)
        fused[eid]["src"].append(f"sem#{rank}")
    for rank, r in enumerate(kw_rows, 1):
        eid = r.get("entry_id") or ""
        # prefer semantic row (carries _distance) if both present
        fused.setdefault(eid, {"row": r, "score": 0.0, "src": []})
        fused[eid]["score"] += 1.0 / (rrf_k + rank)
        fused[eid]["src"].append(f"kw#{rank}")

    ordered = sorted(fused.values(), key=lambda x: x["score"], reverse=True)[:k]
    out = []
    for item in ordered:
        h = _format(item["row"])
        h["score"] = round(item["score"], 5)
        h["sources"] = item["src"]
        out.append(h)
    return out


if __name__ == "__main__":
    mcp.run()
