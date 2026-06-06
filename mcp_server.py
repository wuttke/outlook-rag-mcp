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

import base64
import json
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any

import lancedb
from mcp.server.fastmcp import FastMCP
from sentence_transformers import SentenceTransformer

DB_DIR = Path(os.environ.get("OUTLOOK_RAG_DB", str(Path.home() / "outlook-rag" / "db")))
EXPORT_DIR = Path(os.environ.get(
    "OUTLOOK_RAG_EXPORT_DIR",
    "/mnt/c/Users/wuttke/Documents/outlook-export",
))
ATTACH_PS1 = Path(__file__).resolve().parent / "outlook_attachment.ps1"
DRAFT_PS1 = Path(__file__).resolve().parent / "outlook_draft.ps1"
POWERSHELL_EXE = os.environ.get("OUTLOOK_RAG_POWERSHELL", "powershell.exe")
TABLE_NAME = "messages"
MODEL_NAME = "BAAI/bge-m3"
MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024  # 25 MB safety cap

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
    tbl = ds.scanner(filter=where, columns=PROJECT_COLS).to_table()
    order = "ascending" if sort == "date_asc" else "descending"
    tbl = tbl.sort_by([("date", order)])
    deduped = _dedup_by_email(tbl.to_pylist(), k)
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
    kw_tbl = ds.scanner(filter=where_kw, columns=PROJECT_COLS).to_table()
    kw_tbl = kw_tbl.sort_by([("date", "descending")])
    kw_rows = _dedup_by_email(kw_tbl.to_pylist(), pool)

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


def _wslpath_win(p: str | Path) -> str:
    """Translate a WSL/Linux path to a Windows path usable by powershell.exe."""
    r = subprocess.run(
        ["wslpath", "-w", str(p)],
        capture_output=True, text=True, check=True,
    )
    return r.stdout.strip()


def _run_attachment_ps(action: str, entry_id: str, **kwargs) -> dict[str, Any]:
    """Invoke outlook_attachment.ps1 and parse its single-line JSON stdout."""
    if not ATTACH_PS1.exists():
        return {"error": f"helper script missing: {ATTACH_PS1}"}
    args = [
        POWERSHELL_EXE,
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", _wslpath_win(ATTACH_PS1),
        "-Action", action,
        "-EntryID", entry_id,
    ]
    for k, v in kwargs.items():
        if v is None:
            continue
        args.extend([f"-{k}", str(v)])
    try:
        proc = subprocess.run(
            args, capture_output=True, text=True, timeout=120,
        )
    except FileNotFoundError:
        return {"error": f"powershell.exe not found (set OUTLOOK_RAG_POWERSHELL)"}
    except subprocess.TimeoutExpired:
        return {"error": "powershell helper timed out"}
    out = (proc.stdout or "").strip()
    if not out:
        return {
            "error": "empty response from powershell helper",
            "returncode": proc.returncode,
            "stderr": (proc.stderr or "").strip()[:1000],
        }
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {
            "error": "non-JSON response from powershell helper",
            "returncode": proc.returncode,
            "stdout": out[:1000],
            "stderr": (proc.stderr or "").strip()[:1000],
        }


@mcp.tool()
def list_attachments(entry_id: str) -> dict[str, Any]:
    """List attachments on a mail by talking to the live Outlook session.

    Uses Outlook COM via a PowerShell helper. Inline images (e.g.
    image001.png) and inline embedded items are included, matching what
    Outlook itself shows under "Attachments".

    Args:
        entry_id: Outlook X-Outlook-EntryID (returned by other tools).

    Returns {entry_id, subject, attachments: [{filename, display_name,
    size, type, index}, ...]} or {error: "..."}.
    """
    return _run_attachment_ps("list", entry_id)


@mcp.tool()
def read_attachment(
    entry_id: str,
    filename: str,
    max_bytes: int = MAX_ATTACHMENT_BYTES,
) -> dict[str, Any]:
    """Read one attachment from a mail via the live Outlook session.

    Writes the attachment to a temporary file on the Windows side via
    `Attachment.SaveAsFile`, then reads it back and returns its bytes
    base64-encoded.

    Args:
        entry_id: Outlook X-Outlook-EntryID.
        filename: exact attachment filename (as listed by list_attachments).
        max_bytes: hard cap; larger attachments return a size-only error.

    Returns {filename, display_name, size, type, content_base64} or
    {error: "...", available: [...]}.
    """
    with tempfile.TemporaryDirectory(prefix="outlook-att-") as td:
        # SaveAsFile needs a Windows path; write inside our WSL tempdir and
        # translate. Filename is sanitized to avoid path traversal/quoting.
        safe_name = Path(filename).name or "attachment.bin"
        host_path = Path(td) / safe_name
        win_path = _wslpath_win(host_path)
        result = _run_attachment_ps(
            "read", entry_id, Filename=filename, OutFile=win_path,
        )
        if result.get("error"):
            return result
        try:
            data = host_path.read_bytes()
        except FileNotFoundError:
            return {"error": f"helper wrote no file at {host_path}",
                    "helper_result": result}
        if len(data) > max_bytes:
            return {
                "error": f"attachment too large ({len(data)} > {max_bytes})",
                "filename": result.get("filename"),
                "size": len(data),
                "type": result.get("type"),
            }
        result["size"] = len(data)
        result.pop("out_file", None)
        result["content_base64"] = base64.b64encode(data).decode("ascii")
        return result


def _write_args_file_for_windows(payload: dict[str, Any]) -> tuple[str, str]:
    """Serialise `payload` to a UTF-8 JSON file under the Windows %TEMP%
    directory and return (host_path, win_path). The Windows-side helper is
    responsible for deleting the file once it has read it."""
    win_tmp = subprocess.run(
        [POWERSHELL_EXE, "-NoProfile", "-Command",
         "[Console]::Out.Write($env:TEMP)"],
        capture_output=True, text=True, check=True,
    ).stdout.strip()
    if not win_tmp:
        raise RuntimeError("could not resolve Windows %TEMP%")
    host_tmp = Path(subprocess.run(
        ["wslpath", "-u", win_tmp],
        capture_output=True, text=True, check=True,
    ).stdout.strip())
    import uuid
    name = f"outlook-draft-args-{uuid.uuid4().hex}.json"
    host_path = host_tmp / name
    host_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    win_path = _wslpath_win(host_path)
    return str(host_path), win_path


@mcp.tool()
def create_draft(
    to: str | None = None,
    subject: str | None = None,
    body: str | None = None,
    cc: str | None = None,
    bcc: str | None = None,
    html_body: str | None = None,
    in_reply_to_entry_id: str | None = None,
    reply_all: bool = False,
) -> dict[str, Any]:
    """Create a mail in Outlook's Drafts folder (never sends).

    Either compose a new mail, or — if `in_reply_to_entry_id` is given —
    create a draft reply pre-populated with quoted history. The draft is
    saved (visible in Outlook's Drafts folder) so the user can review and
    send it manually.

    Args:
        to:                 semicolon-separated recipients (or None when replying).
        subject:            mail subject (defaults to the reply subject when replying).
        body:               plain-text body. Ignored if html_body is set.
        cc, bcc:            optional additional recipient lists.
        html_body:          HTML body; takes precedence over `body`.
        in_reply_to_entry_id: when set, create as Reply/ReplyAll to that mail.
        reply_all:          when replying, use ReplyAll instead of Reply.

    Returns {entry_id, subject, to, cc, body_format, draft_folder, ...} or
    {error: "..."}.
    """
    if not DRAFT_PS1.exists():
        return {"error": f"helper script missing: {DRAFT_PS1}"}
    if not any([to, subject, body, html_body, in_reply_to_entry_id]):
        return {"error": "nothing to do: provide at least one of to/subject/body/html_body/in_reply_to_entry_id"}
    payload = {
        "To": to, "Cc": cc, "Bcc": bcc,
        "Subject": subject,
        "Body": body, "HtmlBody": html_body,
        "InReplyToEntryID": in_reply_to_entry_id,
        "ReplyAll": bool(reply_all),
    }
    try:
        _, win_args = _write_args_file_for_windows(payload)
    except Exception as e:
        return {"error": f"failed to stage args file: {e}"}
    args = [
        POWERSHELL_EXE, "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", _wslpath_win(DRAFT_PS1),
        "-ArgsFile", win_args,
    ]
    try:
        proc = subprocess.run(args, capture_output=True, text=True, timeout=60)
    except subprocess.TimeoutExpired:
        return {"error": "powershell helper timed out"}
    out = (proc.stdout or "").strip()
    if not out:
        return {
            "error": "empty response from powershell helper",
            "returncode": proc.returncode,
            "stderr": (proc.stderr or "").strip()[:1000],
        }
    try:
        return json.loads(out)
    except json.JSONDecodeError:
        return {
            "error": "non-JSON response from powershell helper",
            "stdout": out[:1000],
            "stderr": (proc.stderr or "").strip()[:1000],
        }


if __name__ == "__main__":
    mcp.run()
