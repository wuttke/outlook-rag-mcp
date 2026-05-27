"""
Ingest Outlook mbox export into a LanceDB vector store.

Usage:
    python ingest.py [--limit N] [--folders A,B,...] [--batch 32]
"""
from __future__ import annotations

import argparse
import email
import mailbox
import re
import sys
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
from typing import Iterable

import lancedb
import pyarrow as pa
from bs4 import BeautifulSoup
from sentence_transformers import SentenceTransformer
from tqdm import tqdm

EXPORT_DIR = Path("/mnt/c/Users/wuttke/Documents/outlook-export")
DB_DIR = Path.home() / "outlook-rag" / "db"
TABLE_NAME = "messages"
MODEL_NAME = "BAAI/bge-m3"
EMBED_DIM = 1024

# Folder allowlist (human mail only) — match X-Outlook-Folder values
HUMAN_FOLDERS = {
    "Archiv",
    "Archive",
    "Gesendete Elemente",
    "Posteingang",
    "Entwürfe",
}

# Chunking (chars; bge-m3 tokenizer averages ~3.5 chars/token for DE+EN mix)
MAX_CHUNK_CHARS = 2000
OVERLAP_CHARS = 200
MIN_CHUNK_CHARS = 50  # skip near-empty chunks


def decode_part(part) -> str:
    """Decode an email part to text, best-effort."""
    payload = part.get_payload(decode=True)
    if not payload:
        return ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except (LookupError, TypeError):
        return payload.decode("utf-8", errors="replace")


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "head"]):
        tag.decompose()
    text = soup.get_text(separator="\n")
    return text


def extract_body(msg: email.message.Message) -> str:
    """Walk MIME tree, return best text representation."""
    plain_parts: list[str] = []
    html_parts: list[str] = []
    other_parts: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            if part.is_multipart():
                continue
            ct = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                continue
            if ct == "text/plain":
                plain_parts.append(decode_part(part))
            elif ct == "text/html":
                html_parts.append(decode_part(part))
            elif ct.startswith("text/"):
                other_parts.append(decode_part(part))
    else:
        # Non-multipart: PR_TRANSPORT_MESSAGE_HEADERS export often yields TNEF or
        # bare body. get_payload(decode=True) returns text the exporter appended.
        text = decode_part(msg)
        ct = msg.get_content_type()
        if ct == "text/html":
            html_parts.append(text)
        else:
            plain_parts.append(text)

    if plain_parts:
        return "\n\n".join(p for p in plain_parts if p.strip())
    if html_parts:
        return "\n\n".join(html_to_text(p) for p in html_parts if p.strip())
    return "\n\n".join(other_parts).strip()


_WS_RE = re.compile(r"[ \t]+")
_NL_RE = re.compile(r"\n{3,}")
_QUOTE_RE = re.compile(r"^>.*$", re.MULTILINE)


def clean_text(s: str) -> str:
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = _WS_RE.sub(" ", s)
    s = _NL_RE.sub("\n\n", s)
    return s.strip()


def chunk_text(text: str) -> list[str]:
    """Paragraph-aware chunking with overlap."""
    if len(text) <= MAX_CHUNK_CHARS:
        return [text] if len(text) >= MIN_CHUNK_CHARS else []
    paras = [p for p in text.split("\n\n") if p.strip()]
    chunks: list[str] = []
    buf = ""
    for p in paras:
        if len(buf) + len(p) + 2 <= MAX_CHUNK_CHARS:
            buf = f"{buf}\n\n{p}" if buf else p
        else:
            if buf:
                chunks.append(buf)
            if len(p) <= MAX_CHUNK_CHARS:
                # carry overlap from previous chunk
                tail = chunks[-1][-OVERLAP_CHARS:] if chunks else ""
                buf = f"{tail}\n\n{p}" if tail else p
            else:
                # paragraph itself too big: hard-split
                for i in range(0, len(p), MAX_CHUNK_CHARS - OVERLAP_CHARS):
                    chunks.append(p[i : i + MAX_CHUNK_CHARS])
                buf = ""
    if buf and len(buf) >= MIN_CHUNK_CHARS:
        chunks.append(buf)
    return chunks


def hdr(raw) -> str:
    """Coerce a possibly RFC2047-encoded header to a plain str."""
    if raw is None:
        return ""
    if not isinstance(raw, str):
        try:
            raw = str(make_header(decode_header(str(raw))))
        except Exception:
            raw = str(raw)
    return raw


def parse_addresses(raw, limit: int = 200) -> str:
    raw = hdr(raw)
    if not raw:
        return ""
    pairs = getaddresses([raw])
    addrs = [a for _, a in pairs if a]
    return ", ".join(addrs)[:limit]


def parse_date(raw: str | None) -> datetime:
    if not raw:
        return datetime(1970, 1, 1, tzinfo=timezone.utc)
    try:
        dt = parsedate_to_datetime(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return datetime(1970, 1, 1, tzinfo=timezone.utc)


def iter_messages(mbox_path: Path, folder_allowlist: set[str]):
    """Yield (meta_dict, body_text) for each acceptable message."""
    mb = mailbox.mbox(str(mbox_path))
    for msg in mb:
        folder = hdr(msg.get("X-Outlook-Folder")).strip()
        # Exact-match filter: subfolders are NOT included implicitly.
        if folder_allowlist and folder not in folder_allowlist:
            continue
        body = clean_text(extract_body(msg))
        if not body:
            continue
        meta = {
            "entry_id": hdr(msg.get("X-Outlook-EntryID")).strip(),
            "folder": folder,
            "subject": hdr(msg.get("Subject")).strip()[:500],
            "from_addr": parse_addresses(msg.get("From")),
            "to_addr": parse_addresses(msg.get("To"), limit=500),
            "cc_addr": parse_addresses(msg.get("Cc"), limit=500),
            "date": parse_date(hdr(msg.get("Date"))),
            "message_id": hdr(msg.get("Message-ID")).strip()[:500],
            "attachments": hdr(msg.get("X-Outlook-Attachments")).strip()[:1000],
        }
        yield meta, body



SCHEMA = pa.schema([
    pa.field("id", pa.string()),
    pa.field("entry_id", pa.string()),
    pa.field("chunk_idx", pa.int32()),
    pa.field("folder", pa.string()),
    pa.field("subject", pa.string()),
    pa.field("from_addr", pa.string()),
    pa.field("to_addr", pa.string()),
    pa.field("cc_addr", pa.string()),
    pa.field("date", pa.timestamp("us", tz="UTC")),
    pa.field("message_id", pa.string()),
    pa.field("attachments", pa.string()),
    pa.field("body_chunk", pa.string()),
    pa.field("vector", pa.list_(pa.float32(), EMBED_DIM)),
])


def build_records(folder_allowlist: set[str], limit: int | None):
    """Yield record dicts (without vector) from all mboxes."""
    mbox_files = sorted(EXPORT_DIR.glob("*.mbox"))
    total = 0
    for mb_path in mbox_files:
        for meta, body in iter_messages(mb_path, folder_allowlist):
            chunks = chunk_text(body)
            for idx, ch in enumerate(chunks):
                rec = {**meta, "chunk_idx": idx, "body_chunk": ch}
                rec["id"] = f"{meta['entry_id']}#{idx}"
                yield rec
                total += 1
                if limit and total >= limit:
                    return


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None, help="cap on chunks ingested")
    ap.add_argument("--folders", type=str, default=None,
                    help="comma-separated top-level folders (default: human mail)")
    ap.add_argument("--batch", type=int, default=16, help="embed batch size")
    ap.add_argument("--reset", action="store_true", help="drop and recreate table")
    args = ap.parse_args()

    allowlist = HUMAN_FOLDERS if args.folders is None else set(args.folders.split(","))
    print(f"Allowed top-level folders: {sorted(allowlist)}")

    print(f"Loading model {MODEL_NAME} (CPU)...")
    model = SentenceTransformer(MODEL_NAME, device="cpu")
    model.max_seq_length = 512  # enough for 2000-char chunks, much faster than 8192

    DB_DIR.mkdir(parents=True, exist_ok=True)
    db = lancedb.connect(str(DB_DIR))
    if args.reset and TABLE_NAME in db.table_names():
        db.drop_table(TABLE_NAME)
    if TABLE_NAME in db.table_names():
        table = db.open_table(TABLE_NAME)
        existing_ids = {r["id"] for r in table.search().select(["id"]).limit(10**7).to_list()}
        print(f"Existing chunks in table: {len(existing_ids)}")
    else:
        table = db.create_table(TABLE_NAME, schema=SCHEMA)
        existing_ids = set()

    print("Scanning mboxes for new chunks...")
    pending: list[dict] = []
    for rec in build_records(allowlist, args.limit):
        if rec["id"] in existing_ids:
            continue
        pending.append(rec)
    print(f"New chunks to embed: {len(pending)}")
    if not pending:
        return

    print(f"Embedding in batches of {args.batch}...")
    for i in tqdm(range(0, len(pending), args.batch), unit="batch"):
        batch = pending[i : i + args.batch]
        texts = [f"{r['subject']}\n\n{r['body_chunk']}" for r in batch]
        vecs = model.encode(
            texts,
            batch_size=args.batch,
            normalize_embeddings=True,
            show_progress_bar=False,
            convert_to_numpy=True,
        )
        for r, v in zip(batch, vecs):
            r["vector"] = v.tolist()
        table.add(batch)

    print(f"Done. Table rows: {table.count_rows()}")


if __name__ == "__main__":
    main()

