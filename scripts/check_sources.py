"""Check only registered public documentation URLs, and only with --allow-network."""
from __future__ import annotations
import argparse
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from urllib.parse import urlsplit
from common import ROOT, read_json, write_new_json, validate_new_output

ALLOWED_HOSTS = {"learn.microsoft.com", "github.com", "raw.githubusercontent.com", "code.visualstudio.com", "creativecommons.org"}

def validate_url(url: str) -> str:
    p = urlsplit(url)
    if p.scheme != "https" or p.hostname not in ALLOWED_HOSTS or p.username or p.password or p.port not in (None, 443):
        raise ValueError("Only HTTPS URLs on the fixed public documentation allowlist are permitted")
    if p.query or p.fragment:
        raise ValueError("Query strings and fragments are not sent by this checker")
    return url

class RestrictedRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validate_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)

def check_one(url: str, opener=None) -> dict:
    validate_url(url)
    opener = opener or urllib.request.build_opener(RestrictedRedirect())
    request = urllib.request.Request(url, headers={"User-Agent": "ExcelVBAExpertSourceCheck/1.0"}, method="GET")
    try:
        with opener.open(request, timeout=15) as response:
            final = response.geturl()
            validate_url(final)
            response.read(1024)  # No full-page scraper or content cache.
            return {"url": url, "status": "reachable", "http_status": response.status, "final_url": final,
                    "meaning": "Reachability only; not semantic or freshness verification"}
    except (OSError, ValueError, urllib.error.URLError) as exc:
        return {"url": url, "status": "failed", "error": str(exc)}

def run_checks(sources: list[dict], allow_network: bool, limit: int = 10, opener=None) -> dict:
    selected = sources[:limit]
    for d in selected:
        validate_url(d["url"])
    if not allow_network:
        return {"status": "DRY_RUN_NO_NETWORK", "planned_sources": [d["id"] for d in selected],
                "instruction": "Use --allow-network only after explicit user permission"}
    results = []
    for d in selected:
        results.append(dict(check_one(d["url"], opener), id=d["id"]))
        if opener is None:
            time.sleep(0.25)
    return {"status": "NETWORK_CHECK_EXECUTED", "results": results}

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--allow-network", action="store_true")
    p.add_argument("--limit", type=int, default=10)
    p.add_argument("--out", type=Path)
    a = p.parse_args()
    if not 1 <= a.limit <= 100:
        p.error("--limit must be between 1 and 100")
    try:
        validate_new_output(a.out)
        result = run_checks(read_json(ROOT / "metadata/source-manifest.json")["sources"], a.allow_network, a.limit)
        if a.out:
            write_new_json(a.out, result)
        print(json.dumps(result, indent=2))
        return 1 if any(d["status"] == "failed" for d in result.get("results", [])) else 0
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
