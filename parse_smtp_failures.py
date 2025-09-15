#!/usr/bin/env python3
"""
parse_smtp_failures.py

Parse lines like:
  2025-09-15 10:42:13 | INFO | root | {"email": "...", "ok": false, "error": "...\\n..."}
and emit a list of failures with a concise reason.

Usage:
  python parse_smtp_failures.py --in logfile.txt --format json
  python parse_smtp_failures.py --in logfile.txt --format csv

If --in is omitted, reads from stdin.
"""

import argparse
import csv
import io
import json
import re
import sys

# Heuristic: extract a concise reason from a libcurl SMTP transcript
def summarize_error(err_text: str) -> str:
    if not err_text:
        return "Unknown error"

    # Convert escaped newlines to actual newlines if the string came from JSON
    # (json.loads already does this; this is safe regardless)
    text = err_text

    # Split into lines
    lines = text.splitlines()

    # Prefer the last server reply line (starts with '< ')
    server_lines = [ln for ln in lines if ln.startswith("< ")]
    if server_lines:
        last = server_lines[-1][2:].strip()  # strip "< "
        # If it looks like a 5xx SMTP error, keep it as-is
        m = re.match(r"(\d{3}[\s-].*)", last)
        if m:
            return m.group(1)
        return last

    # Next, prefer the last libcurl status line (starts with '* ')
    star_lines = [ln for ln in lines if ln.startswith("* ")]
    if star_lines:
        return star_lines[-1][2:].strip()

    # Next, the last client command (starts with '> ')
    client_lines = [ln for ln in lines if ln.startswith("> ")]
    if client_lines:
        # Hide long base64 XOAUTH2 blobs
        last = client_lines[-1][2:].strip()
        if last.startswith("dXNlcj0") or "Bearer " in last:
            return "Sent XOAUTH2 blob"
        return last

    # Fallback: first line or first 140 chars
    trimmed = lines[0].strip() if lines else text.strip()
    return (trimmed[:140] + "…") if len(trimmed) > 140 else trimmed


def parse_log_line(line: str):
    """
    Split 'timestamp | level | logger | {json}' safely.
    We assume the JSON blob is the part after the final ' | ' sequence.
    """
    # Fast path: split into four chunks max
    parts = line.split(" | ", 3)
    if len(parts) < 4:
        return None  # not our format
    ts, level, logger, json_str = parts[0], parts[1], parts[2], parts[3].strip()

    # Some logs can contain trailing spaces or extra text; ensure JSON starts at first '{'
    start = json_str.find("{")
    if start > 0:
        json_str = json_str[start:]

    try:
        payload = json.loads(json_str)
    except json.JSONDecodeError:
        return None

    return {
        "timestamp": ts.strip(),
        "level": level.strip(),
        "logger": logger.strip(),
        "payload": payload,
    }


def main():
    ap = argparse.ArgumentParser(description="Parse SMTP failure entries from logs.")
    ap.add_argument("--in", dest="infile", help="Input log file (default: stdin)")
    ap.add_argument("--format", choices=["json", "csv"], default="json", help="Output format")
    args = ap.parse_args()

    if args.infile:
        fh = open(args.infile, "r", encoding="utf-8", errors="replace")
    else:
        fh = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8", errors="replace")

    log_entry_count = 0

    failures = []
    for raw in fh:
        raw = raw.rstrip("\n")
        if not raw.strip():
            continue
        parsed = parse_log_line(raw)

        log_entry_count = log_entry_count + 1

        if not parsed:
            continue

        pl = parsed["payload"]
        ok = pl.get("ok", None)
        if ok is True:
            continue  # success; skip
        # Treat missing 'ok' as failure conservatively
        if ok is False or ok is None:
            email = pl.get("email") or pl.get("to") or ""
            firstname = pl.get("firstname", "")
            lastname = pl.get("lastname", "")
            country = pl.get("country", "")
            error = pl.get("error", "")

            failures.append({
                "timestamp": parsed["timestamp"],
                "email": email,
                "firstname": firstname,
                "lastname": lastname,
                "country": country,
                "reason": summarize_error(error),
            })

    if args.infile:
        fh.close()

    if args.format == "json":
        json.dump(failures, sys.stdout, indent=2, ensure_ascii=False)
        sys.stdout.write("\n")
    else:
        writer = csv.DictWriter(sys.stdout,
                                fieldnames=["timestamp", "email", "firstname", "lastname", "country", "reason"])
        writer.writeheader()
        for row in failures:
            writer.writerow(row)


    print( f"Parsed {log_entry_count} log entries.\n" )


if __name__ == "__main__":
    main()
