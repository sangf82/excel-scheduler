#!/usr/bin/env python3
"""
One-time setup script to create public GitHub Gists for install.ps1 and install.sh.
After creation, add the printed Gist IDs and your PAT to the repository secrets.

Requires: Python 3.8+ (uses only stdlib)
"""

import json
import os
import sys
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError


def create_gist(filename: str, content: str, description: str, pat: str) -> dict:
    """Create a public GitHub Gist and return the JSON response."""
    payload = {
        "description": description,
        "public": True,
        "files": {
            filename: {"content": content}
        }
    }
    data = json.dumps(payload).encode("utf-8")
    req = Request(
        "https://api.github.com/gists",
        data=data,
        headers={
            "Authorization": f"token {pat}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urlopen(req) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except HTTPError as e:
        body = e.read().decode("utf-8")
        print(f"\n[ERROR] GitHub API returned {e.code}: {body}", file=sys.stderr)
        sys.exit(1)


def main():
    project_root = Path(__file__).parent.parent.resolve()
    ps1_path = project_root / "scripts" / "install.ps1"
    sh_path = project_root / "scripts" / "install.sh"

    if not ps1_path.exists():
        print(f"[ERROR] Not found: {ps1_path}", file=sys.stderr)
        sys.exit(1)
    if not sh_path.exists():
        print(f"[ERROR] Not found: {sh_path}", file=sys.stderr)
        sys.exit(1)

    print("=" * 60)
    print("MedMate Scheduler — Public Gist Setup")
    print("=" * 60)
    print()
    print("This script creates two PUBLIC GitHub Gists for your install")
    print("scripts so they remain accessible after you make the repo PRIVATE.")
    print()
    print("You need a GitHub Personal Access Token (classic) with 'gist' scope.")
    print("Create one at: https://github.com/settings/tokens/new")
    print()

    pat = input("Paste your GitHub PAT (hidden): ").strip()
    if not pat:
        print("[ERROR] PAT is required.", file=sys.stderr)
        sys.exit(1)

    print("\nCreating Gist for install.ps1 ...")
    ps1_content = ps1_path.read_text(encoding="utf-8")
    ps1_gist = create_gist(
        "install.ps1",
        ps1_content,
        "MedMate Scheduler — Windows install script (auto-synced)",
        pat,
    )
    ps1_id = ps1_gist["id"]
    ps1_raw = ps1_gist["files"]["install.ps1"]["raw_url"]

    print("Creating Gist for install.sh ...")
    sh_content = sh_path.read_text(encoding="utf-8")
    sh_gist = create_gist(
        "install.sh",
        sh_content,
        "MedMate Scheduler — macOS install script (auto-synced)",
        pat,
    )
    sh_id = sh_gist["id"]
    sh_raw = sh_gist["files"]["install.sh"]["raw_url"]

    print("\n" + "=" * 60)
    print("SUCCESS! Gists created.")
    print("=" * 60)
    print()
    print("Copy the following values into your repository secrets:")
    print()
    print("  Repository -> Settings -> Secrets and variables -> Actions")
    print()
    print(f"  GIST_PAT     = {pat[:4]}{'*' * (len(pat) - 8)}{pat[-4:]}")
    print(f"  GIST_ID_PS1  = {ps1_id}")
    print(f"  GIST_ID_SH   = {sh_id}")
    print()
    print("Update your README remote-install URLs to:")
    print()
    print(f"  Windows:  irm {ps1_raw} | iex")
    print(f"  macOS:    curl -fsSL {sh_raw} | bash")
    print()
    print("After the secrets are saved, every push to 'main' that changes")
    print("install.ps1 or install.sh will auto-sync via GitHub Actions.")
    print()
    print("=" * 60)


if __name__ == "__main__":
    main()
