#!/usr/bin/env python3
"""
One-command pipeline: install deps, run conversion, verify output.

This wraps check_setup.py and convert_dashboard_claude.py.
"""

import argparse
import subprocess
import sys
from pathlib import Path
import os


def _run(cmd: list[str]) -> int:
    result = subprocess.run(cmd)
    return result.returncode


def _is_pbix_or_pbip(source: str) -> bool:
    path = Path(source)
    if path.is_dir():
        return any(p.suffix.lower() == ".pbip" for p in path.glob("*.pbip"))
    return path.suffix.lower() in {".pbip", ".pbix"}


def _resolve_assistant(selection: str) -> str:
    if selection in ("claude", "copilot"):
        return selection

    claude_markers = [
        "CLAUDE_CODE",
        "ANTHROPIC_API_KEY",
        "CLAUDE_SESSION",
    ]
    for key in claude_markers:
        if os.environ.get(key):
            return "claude"
    return "copilot"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Install deps, run conversion, and validate output."
    )
    ap.add_argument("--source", required=True, help="Input PPTX/PDF/PBIP/PBIX path")
    ap.add_argument("--assistant", choices=["claude", "copilot", "auto"], default="auto")
    ap.add_argument("--profile", choices=["claude", "copilot"], default=None)
    ap.add_argument("--output", default=None, help="Output PPTX path")
    ap.add_argument("--context", default=None, help="Optional analysis focus")
    args = ap.parse_args()

    resolved_assistant = _resolve_assistant(args.assistant)
    profile = args.profile or resolved_assistant

    print("=" * 70)
    print("STEP 0: INSTALLING DEPENDENCIES")
    print("=" * 70)
    setup_cmd = [
        sys.executable,
        "check_setup.py",
        "--profile",
        profile,
        "--auto-install",
    ]
    rc = _run(setup_cmd)
    if rc != 0:
        return rc

    if _is_pbix_or_pbip(args.source):
        print("\n" + "=" * 70)
        print("MCP CHECK")
        print("=" * 70)
        check_cmd = [
            sys.executable,
            "setup_pbi_mcp.py",
            "--check",
        ]
        mcp_rc = _run(check_cmd)
        if mcp_rc != 0:
            print("\nWARNING: MCP not ready. PBIP/PBIX will fall back to image-only analysis.\n")

    print("\n" + "=" * 70)
    print("STEP 1-3: CONVERSION")
    print("=" * 70)

    convert_cmd = [
        sys.executable,
        "convert_dashboard_claude.py",
        "--source",
        args.source,
        "--assistant",
        args.assistant,
    ]
    if args.output:
        convert_cmd.extend(["--output", args.output])
    if args.context:
        convert_cmd.extend(["--context", args.context])

    return _run(convert_cmd)


if __name__ == "__main__":
    raise SystemExit(main())
