#!/usr/bin/env python3

"""
parse_recipient_info.py

This script parses the input and emits values appropriate for consumption by the sender script.

"""


import argparse
import csv
import json
import os
import sys
import logging
from typing import List, Dict, Any, Optional, Union
from datetime import datetime, timezone, timedelta

from typing import TypedDict, Literal, Union

import json, time, random, subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed

from m365_oauth_tokeninfo import OAuthToken


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(threadName)s %(levelname)s %(message)s",
    stream=sys.stdout,
)

log = logging.getLogger("mailmerge")


# -------- CSV --------
def read_csv_to_rows(path: str) -> List[Dict[str, Any]]:
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return [dict(row) for row in reader]

# -------- Excel (.xlsx / .xlsm) --------
def read_excel_to_rows(path: str, sheet: Optional[str] = None) -> List[Dict[str, Any]]:
    try:
        from openpyxl import load_workbook  # type: ignore
    except Exception as e:
        raise RuntimeError(
            "Reading Excel files requires the 'openpyxl' package. Install with: pip install openpyxl"
        ) from e

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb.active

    rows_iter = ws.iter_rows(values_only=True)
    try:
        headers = next(rows_iter)
    except StopIteration:
        return []

    headers = [str(h) if h is not None else "" for h in headers]
    out: List[Dict[str, Any]] = []
    for row in rows_iter:
        record = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
        out.append(record)
    return out

def to_keyed_dict(rows: List[Dict[str, Any]], key_column: str) -> Dict[str, Dict[str, Any]]:
    mapping: Dict[str, Dict[str, Any]] = {}
    for i, r in enumerate(rows, start=1):
        key = r.get(key_column)
        if key is None:
            # Normalize None to empty string to avoid TypeError on dict keys
            key = ""
        key = str(key)
        mapping[key] = r  # last one wins on duplicate keys
    return mapping

def infer_file_type(path: str) -> Optional[str]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return "csv"
    if ext in {".xlsx", ".xlsm"}:
        return "excel"
    if ext == ".xls":
        return "xls"  # not supported here
    return None


def token_expired(token: Dict, skew_seconds: int = 600) -> bool:
    """
    Returns True if the token is expired or will expire within `skew_seconds`.
    Expects JSON fields: obtained_at (unix seconds) and expires_in (seconds).
    """
    
    return token.seconds_until_expiry < skew_seconds



MAX_CONCURRENCY = 12            # threads in parallel (tune carefully)
RATE_PER_MINUTE = 600           # hard cap across all threads
SPACING = 60.0 / max(1, RATE_PER_MINUTE)


class SendSuccess(TypedDict):
    email: str
    ok: Literal[True]
    ms: int

class SendError(TypedDict):
    email: str
    ok: Literal[False]
    error: str

SendResult = Union[SendSuccess, SendError]



def send_one(tokenstr: str, fromname:str, fromaddr: str, subject: str, emailhtmlpath: str, rec: Dict) -> SendResult:
    # light jitter to avoid thundering herd
    time.sleep(random.uniform(0, SPACING))

    #DEBUG
    # print( f"Going to send e-mail to: {rec['firstname']} {rec['lastname']} to {rec['EmailsToUse']}" )
    log.info( f"Going to send e-mail to: {rec['firstname']} {rec['lastname']} to {rec['EmailsToUse']}" )

    toAddr = "mikecress+wafuniftest@gmail.com"

    # build your CLI invocation (example args)
    # cmd = [
    #     "build/email-sender",                      # your email CLI
    #     "--to", rec["email"],
    #     "--subject", rec["subject"],
    #     "--body", rec["body"],               # or --body-file path
    # ]
    cmd = [
        "build/email-sender",
        "--from_name", fromname,
        "--from", fromaddr,
        "--to", toAddr,
        "--subject", subject,
        "--username", "l.bucur.cress@wafunif.org",
        "--file", emailhtmlpath,
        "--token", tokenstr
    ]

    backoff = 1.0
    for attempt in range(1, 6):  # up to 5 tries
        t0 = time.time()
        p = subprocess.run(cmd, capture_output=True, text=True)
        dt = time.time() - t0

        if p.returncode == 0:
            return {"email": toAddr, "ok": True, "ms": int(dt*1000)}
        # transient? back off and retry
        time.sleep(backoff + random.uniform(0, 0.200))
        backoff = min(backoff * 2, 16.0)

    return {
        "email": toAddr,
        "ok": False,
        "error": (p.stderr or p.stdout).strip()[:2000]
    }


def run_mail_merge(token, fromname, fromaddr, subject, emailhtmlpath, records):
    results = []
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as pool:
        futures = [pool.submit(send_one, token.access_token, fromname, fromaddr, subject, emailhtmlpath, records[0]) ]
        # futures = [pool.submit(send_one, token.access_token, fromaddr, subject, emailhtmlpath, r) for r in records]
        for fut in as_completed(futures):
            res = fut.result()
            print(json.dumps(res, ensure_ascii=False))
            results.append(res)
    return results

def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Read a CSV or Excel spreadsheet into dict structures.")
    parser.add_argument("path", help="Path to the CSV or Excel file.")
    parser.add_argument("--tokenfile", help="File containing the OAuth token for authentication to the Microsoft 365 SMTP server. (Obtained by running m365_token_helper.py)")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--type", choices=["csv", "excel"], help="Explicitly set the file type.")
    parser.add_argument("--sheet", help="Excel sheet name (defaults to the first/active sheet).", default=None)
    parser.add_argument("--output", "-o", help="If set, write program output to this path.", default=None)
    args = parser.parse_args(argv)

    if args.tokenfile is None:
        print("Please supply a valid token file value.", file=sys.stderr)
        return 2

    if args.type is None:
        print("Please pass --type {csv,excel} or --excel.", file=sys.stderr)
        return 2

    if args.type == "excel" and infer_file_type(args.path) == "xls":
        print(".xls is not supported by this script. Save as .xlsx or .csv, or install a library that supports .xls.", file=sys.stderr)
        return 2

    try:
        with open(f"{args.tokenfile}", "r", encoding="utf-8") as f:
            token_data = json.load(f)
    except FileNotFoundError:
        print(f"{args.tokenfile} not found")
        return 2
    except IsADirectoryError:
        print(f"{args.tokenfile} is a directory, not a file")
        return 2
    except PermissionError:
        print(f"no permission to read {args.tokenfile}")
        return 2

    
    tok = OAuthToken(**token_data)

    # print(f"Token data is {tok}\n")

    has_token_expired = token_expired(tok)

    if has_token_expired:
        print("OAuth bearer token has expired. Please refresh it and then retry this script. See README.txt for instructions on how to do this.\n")
        return 2
    

    try:
        if args.type == "csv":
            rows = read_csv_to_rows(args.path)
        elif args.type == "excel":
            rows = read_excel_to_rows(args.path, sheet=args.sheet)
        else:
            print(f"Unsupported file type: {args.type}", file=sys.stderr)
            return 2
    except Exception as e:
        print(f"Error reading file: {e}", file=sys.stderr)
        return 1

    # for row in rows:
        # print( "", row['firstname'], row['lastname'], row['EmailsToUse'] )

    subject = "howdy"
    fromname = "Big Michael"
    fromaddr = "membership@wafunif.org"
    emailhtmlpath = "email.html"
    
    run_mail_merge( tok, fromname, fromaddr, subject, emailhtmlpath, rows )


    return 0

if __name__ == "__main__":
    raise SystemExit(main())
