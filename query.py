"""
Query the Outlook RAG vector store.

Usage:
    python query.py "your question or topic"
    python query.py --folder Archiv --since 2025-01-01 -k 5 "topic"
"""
from __future__ import annotations

import argparse
import shutil
import sys
import textwrap
from datetime import datetime, timezone
from pathlib import Path

import lancedb
from sentence_transformers import SentenceTransformer

DB_DIR = Path.home() / "outlook-rag" / "db"
TABLE_NAME = "messages"
MODEL_NAME = "BAAI/bge-m3"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("query", nargs="+", help="search query (natural language)")
    ap.add_argument("-k", type=int, default=8, help="results to return")
    ap.add_argument("--folder", type=str, default=None, help="restrict to folder (exact)")
    ap.add_argument("--from-addr", type=str, default=None, help="substring of sender")
    ap.add_argument("--since", type=str, default=None, help="ISO date, e.g. 2024-01-01")
    ap.add_argument("--until", type=str, default=None, help="ISO date")
    ap.add_argument("--no-truncate", action="store_true", help="show full body chunk")
    args = ap.parse_args()

    q = " ".join(args.query)
    print(f"Query: {q}\n")

    model = SentenceTransformer(MODEL_NAME, device="cpu")
    model.max_seq_length = 512
    qvec = model.encode([q], normalize_embeddings=True)[0].tolist()

    db = lancedb.connect(str(DB_DIR))
    table = db.open_table(TABLE_NAME)

    builder = table.search(qvec).metric("cosine")
    filters: list[str] = []
    if args.folder:
        filters.append(f"folder = '{args.folder}'")
    if args.from_addr:
        filters.append(f"from_addr LIKE '%{args.from_addr}%'")
    if args.since:
        filters.append(f"date >= timestamp '{args.since}T00:00:00Z'")
    if args.until:
        filters.append(f"date <= timestamp '{args.until}T23:59:59Z'")
    if filters:
        where = " AND ".join(filters)
        print(f"Filter: {where}\n")
        builder = builder.where(where, prefilter=True)

    rows = builder.limit(args.k).to_list()

    width = shutil.get_terminal_size((100, 24)).columns
    body_w = max(60, width - 4)
    for i, r in enumerate(rows, 1):
        score = 1.0 - r.get("_distance", 0.0)  # cosine distance -> similarity
        date = r["date"].strftime("%Y-%m-%d %H:%M") if r.get("date") else "????"
        chunk_marker = f" [chunk {r['chunk_idx']}]" if r["chunk_idx"] else ""
        print(f"#{i}  score={score:.3f}  [{r['folder']}]  {date}{chunk_marker}")
        print(f"    From:    {r['from_addr']}")
        if r["to_addr"]:
            print(f"    To:      {r['to_addr'][:body_w-13]}")
        print(f"    Subject: {r['subject']}")
        body = r["body_chunk"].strip()
        if not args.no_truncate and len(body) > 400:
            body = body[:400] + " …"
        for line in textwrap.wrap(body, body_w, replace_whitespace=False,
                                  drop_whitespace=False):
            print(f"    {line}")
        print()


if __name__ == "__main__":
    main()
