#!/usr/bin/env python3
"""
m365_token_helper.py

Obtain and refresh OAuth2 tokens for Microsoft 365 SMTP (XOAUTH2) using the Device Code flow.

Requirements:
  - Python 3.7+
  - requests (pip install requests)

Typical usage:
  # 1) Interactive login with Device Code (prints URL & code to enter)
  python m365_token_helper.py login \
      --tenant TENANT_ID \
      --client-id CLIENT_ID \
      --scopes "https://outlook.office365.com/SMTP.Send offline_access" \
      --out tokens.json

  # 2) Later, refresh the token
  python m365_token_helper.py refresh \
      --tenant TENANT_ID \
      --client-id CLIENT_ID \
      --in tokens.json \
      --out tokens.json

The resulting access_token can be passed to your C program with --token.
"""

import argparse
import json
import sys
import time
from pathlib import Path

import requests

DEVICE_CODE_URL_TMPL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
TOKEN_URL_TMPL       = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

DEFAULT_SCOPES = "https://outlook.office365.com/SMTP.Send offline_access"


def save_tokens(path: Path, data: dict) -> None:
    # Only store what's useful
    keep = {
        "access_token": data.get("access_token"),
        "refresh_token": data.get("refresh_token"),
        "expires_in": data.get("expires_in"),
        "ext_expires_in": data.get("ext_expires_in"),
        "token_type": data.get("token_type"),
        "scope": data.get("scope"),
        "obtained_at": int(time.time()),
    }
    path.write_text(json.dumps(keep, indent=2))
    print(f"Saved tokens to {path}")


def read_tokens(path: Path) -> dict:
    return json.loads(path.read_text())


def pretty_expiry(obtained_at: int, expires_in: int) -> str:
    if not obtained_at or not expires_in:
        return "unknown"
    ttl = obtained_at + expires_in - int(time.time())
    return f"{ttl} seconds remaining" if ttl > 0 else f"expired {-ttl} seconds ago"


def cmd_login(args):
    tenant = args.tenant
    client_id = args.client_id
    scopes = args.scopes or DEFAULT_SCOPES
    out_path = Path(args.out)

    device_url = DEVICE_CODE_URL_TMPL.format(tenant=tenant)
    token_url  = TOKEN_URL_TMPL.format(tenant=tenant)

    # 1) Start device code flow
    resp = requests.post(
        device_url,
        data={"client_id": client_id, "scope": scopes},
        timeout=30,
    )
    if resp.status_code != 200:
        print("Device code request failed:", resp.status_code, resp.text, file=sys.stderr)
        sys.exit(1)

    dc = resp.json()
    user_code = dc["user_code"]
    verify_uri = dc["verification_uri"] if "verification_uri" in dc else dc["verification_uri_complete"]
    interval = int(dc.get("interval", 5))
    expires_in = int(dc.get("expires_in", 900))

    print("\n=== Device Code ===")
    print(f"Go to: {verify_uri}")
    print(f"Enter code: {user_code}")
    print(f"(Code expires in ~{expires_in} seconds)\n")

    # 2) Poll for token
    start = time.time()
    while True:
        time.sleep(interval)
        resp = requests.post(
            token_url,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id,
                "device_code": dc["device_code"],
            },
            timeout=30,
        )
        if resp.status_code == 200:
            tokens = resp.json()
            save_tokens(out_path, tokens)
            print("\nLogin successful.")
            print(f"Access token type: {tokens.get('token_type')}")
            print(f"Scopes: {tokens.get('scope')}")
            print(f"Access token TTL: {tokens.get('expires_in')} seconds")
            return
        else:
            data = resp.json()
            error = data.get("error")
            if error in ("authorization_pending", "slow_down"):
                # keep polling
                if error == "slow_down":
                    interval += 2
                if time.time() - start > expires_in + 30:
                    print("Device code expired; please run login again.", file=sys.stderr)
                    sys.exit(1)
                continue
            else:
                print("Token request failed:", data, file=sys.stderr)
                sys.exit(1)


def cmd_refresh(args):
    tenant = args.tenant
    client_id = args.client_id
    in_path = Path(args.infile)
    out_path = Path(args.out)

    token_url = TOKEN_URL_TMPL.format(tenant=tenant)

    current = read_tokens(in_path)
    refresh_token = current.get("refresh_token")
    if not refresh_token:
        print("No refresh_token found in input file. Run 'login' first.", file=sys.stderr)
        sys.exit(1)

    resp = requests.post(
        token_url,
        data={
            "grant_type": "refresh_token",
            "client_id": client_id,
            "refresh_token": refresh_token,
        },
        timeout=30,
    )
    if resp.status_code != 200:
        print("Refresh failed:", resp.status_code, resp.text, file=sys.stderr)
        sys.exit(1)

    tokens = resp.json()
    save_tokens(out_path, tokens)
    print("\nRefreshed access token.")
    print(f"Access token TTL: {tokens.get('expires_in')} seconds")


def cmd_show(args):
    data = read_tokens(Path(args.infile))
    obtained_at = data.get("obtained_at")
    expires_in = data.get("expires_in")
    print(json.dumps(data, indent=2))
    print(f"\nStatus: {pretty_expiry(obtained_at, expires_in)}")


def main():
    p = argparse.ArgumentParser(description="Helper for Microsoft 365 SMTP OAuth tokens (Device Code flow).")
    sub = p.add_subparsers(dest="cmd", required=True)

    # login
    pl = sub.add_parser("login", help="Interactive device-code login to obtain tokens")
    pl.add_argument("--tenant", required=True, help="Tenant ID or domain (e.g. contoso.onmicrosoft.com or GUID)")
    pl.add_argument("--client-id", required=True, help="Azure App (client) ID")
    pl.add_argument("--scopes", default=DEFAULT_SCOPES,
                    help=f"Space-separated scopes (default: '{DEFAULT_SCOPES}')")
    pl.add_argument("--out", default="tokens.json", help="Output token file (default: tokens.json)")
    pl.set_defaults(func=cmd_login)

    # refresh
    pr = sub.add_parser("refresh", help="Use refresh_token to get a new access_token")
    pr.add_argument("--tenant", required=True, help="Tenant ID or domain")
    pr.add_argument("--client-id", required=True, help="Azure App (client) ID")
    pr.add_argument("--in", dest="infile", required=True, help="Existing token file (from login)")
    pr.add_argument("--out", default="tokens.json", help="Output token file (default: tokens.json)")
    pr.set_defaults(func=cmd_refresh)

    # show
    ps = sub.add_parser("show", help="Display token file and expiry status")
    ps.add_argument("--in", dest="infile", required=True, help="Token file to inspect")
    ps.set_defaults(func=cmd_show)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
