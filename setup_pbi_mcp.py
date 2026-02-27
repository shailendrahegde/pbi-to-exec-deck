#!/usr/bin/env python3
"""
Power BI Modeling MCP — Setup & Configuration

Downloads, installs, and configures the Microsoft Power BI Modeling MCP server
so Claude Code can query live Power BI Desktop models via DAX.

  https://github.com/microsoft/powerbi-modeling-mcp

Usage:
    python setup_pbi_mcp.py           # auto-detect, install, configure
    python setup_pbi_mcp.py --check   # status check only (no install)
    python setup_pbi_mcp.py --force   # reinstall even if already configured
"""

import argparse
import json
import os
import sys
import zipfile
from pathlib import Path

# ─── Constants ────────────────────────────────────────────────────────────────
MCP_PUBLISHER = "analysis-services"
MCP_EXTENSION = "powerbi-modeling-mcp"
MCP_NAME      = "powerbi-modeling"          # key used inside .mcp.json
MCP_EXE       = "powerbi-modeling-mcp.exe"
MANUAL_DIR    = Path(r"C:\MCPServers\PowerBIModelingMCP")
MCP_JSON      = Path(".mcp.json")


# ─── Pretty printing ──────────────────────────────────────────────────────────
def _banner(text):
    print("\n" + "=" * 68)
    print(f"  {text}")
    print("=" * 68)

def _step(label):  print(f"\n  {label}")
def _ok(msg):      print(f"    OK  {msg}")
def _warn(msg):    print(f"    !!  {msg}")
def _fail(msg):    print(f"    XX  {msg}")


# ─── 1. Detection ─────────────────────────────────────────────────────────────
def find_installed_exe() -> "Path | None":
    """Scan known locations for an already-installed MCP executable."""
    candidates = []

    # a) VS Code extension folder  (~/.vscode/extensions/<publisher>.<ext>-<ver>/)
    vscode_ext = Path(os.environ.get("USERPROFILE", Path.home())) / ".vscode" / "extensions"
    if vscode_ext.exists():
        for d in vscode_ext.glob(f"{MCP_PUBLISHER}.{MCP_EXTENSION}-*"):
            candidates.append(d / "extension" / "server" / MCP_EXE)

    # b) Manual install directory
    candidates.append(MANUAL_DIR / "extension" / "server" / MCP_EXE)

    for c in candidates:
        if c.exists():
            return c.resolve()
    return None


def read_mcp_json_config() -> "tuple[bool, Path | None]":
    """Return (is_configured, verified_exe_path_or_None) from .mcp.json."""
    if not MCP_JSON.exists():
        return False, None
    try:
        cfg    = json.loads(MCP_JSON.read_text(encoding="utf-8"))
        server = cfg.get("mcpServers", {}).get(MCP_NAME)
        if server:
            exe = Path(server.get("command", ""))
            return True, (exe if exe.exists() else None)
    except Exception:
        pass
    return False, None


# ─── 2. Installation ──────────────────────────────────────────────────────────
def _fetch_latest_version() -> str:
    """Query VS Marketplace REST API for the latest extension version."""
    import urllib.request
    url  = "https://marketplace.visualstudio.com/_apis/public/gallery/extensionquery"
    body = json.dumps({
        "filters": [{"criteria": [{"filterType": 7,
                                    "value": f"{MCP_PUBLISHER}.{MCP_EXTENSION}"}]}],
        "flags": 0x100   # IncludeLatestVersionOnly
    }).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/json",
        "Accept":        "application/json;api-version=7.1-preview.1"
    })
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read())
    return data["results"][0]["extensions"][0]["versions"][0]["version"]


def download_and_extract_vsix() -> "Path | None":
    """Download the VSIX package from VS Marketplace and extract to MANUAL_DIR."""
    import urllib.request

    _step("Querying VS Marketplace for latest version...")
    try:
        version = _fetch_latest_version()
        _ok(f"Latest version: {version}")
    except Exception as e:
        version = "0.1.9"
        _warn(f"Marketplace query failed ({e}) — trying version {version}")

    vsix_url = (
        f"https://marketplace.visualstudio.com/_apis/public/gallery/publishers/"
        f"{MCP_PUBLISHER}/vsextensions/{MCP_EXTENSION}/{version}/vspackage"
        f"?targetPlatform=win32-x64"
    )

    Path("temp").mkdir(exist_ok=True)
    tmp = Path("temp") / f"{MCP_EXTENSION}-{version}.vsix"

    _step(f"Downloading package...")
    print(f"    URL: {vsix_url}")
    try:
        def _progress(block_count, block_size, total):
            if total > 0:
                pct = min(block_count * block_size * 100 // total, 100)
                print(f"\r    {pct}%  ({min(block_count * block_size, total) // 1024} / {total // 1024} KB)", end="")
        urllib.request.urlretrieve(vsix_url, tmp, reporthook=_progress)
        print()
        _ok(f"Saved to {tmp}")
    except Exception as e:
        _fail(f"Download failed: {e}")
        return None

    _step(f"Extracting to {MANUAL_DIR} ...")
    MANUAL_DIR.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(tmp, "r") as z:
            z.extractall(MANUAL_DIR)
        tmp.unlink(missing_ok=True)
        _ok("Extracted successfully")
    except Exception as e:
        _fail(f"Extraction failed: {e}")
        return None

    exe = MANUAL_DIR / "extension" / "server" / MCP_EXE
    if exe.exists():
        return exe.resolve()
    _fail(f"Executable not found after extraction (expected: {exe})")
    return None


# ─── 3. Configure .mcp.json ───────────────────────────────────────────────────
def write_mcp_json(exe_path: Path):
    """Merge the powerbi-modeling server entry into .mcp.json."""
    cfg = {}
    if MCP_JSON.exists():
        try:
            cfg = json.loads(MCP_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass

    cfg.setdefault("mcpServers", {})[MCP_NAME] = {
        "command": str(exe_path),
        "args":    ["--start", "--readonly", "--skipconfirmation"]
    }
    MCP_JSON.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    _ok(f"Written: {MCP_JSON}")
    _ok(f"Command: {exe_path}")


# ─── 4. User-facing messages ──────────────────────────────────────────────────
_BOX_W = 66

def _box(*lines):
    border = "-" * _BOX_W
    print(f"+{border}+")
    for line in lines:
        print(f"|  {line:<{_BOX_W - 2}}|")
    print(f"+{border}+")


def print_already_configured(exe_path: Path):
    print()
    _box(
        "Power BI MCP is already installed and configured.",
        "",
        f"Exe:  {str(exe_path)[:_BOX_W - 6]}",
        "",
        "To run PBIP analysis:",
        "  1. Open Power BI Desktop with your .pbip report",
        "  2. Claude Code automatically connects via MCP at startup",
        "  3. python convert_dashboard_claude.py --source report.pbip",
        "",
        "Run with --force to reinstall."
    )


def print_next_steps():
    print()
    _box(
        "SETUP COMPLETE",
        "",
        "Two things to do before running PBIP analysis:",
        "",
        "  1. Open Power BI Desktop and load your .pbip report",
        "     (the MCP connects to the running Desktop process)",
        "",
        "  2. Restart Claude Code to pick up the new MCP server",
        "     Close and relaunch, or: /exit then re-open",
        "",
        "Then run:",
        "  python convert_dashboard_claude.py --source report.pbip",
        "",
        "Claude will query exact DAX values directly from your model."
    )


def print_manual_steps():
    _box(
        "Automatic installation failed. Manual steps:",
        "",
        "1. Open this URL in your browser:",
        "   https://marketplace.visualstudio.com/items",
        "   ?itemName=analysis-services.powerbi-modeling-mcp",
        "",
        "2. Click 'Download Extension', rename the .vsix to .zip",
        "",
        "3. Extract to:  C:\\MCPServers\\PowerBIModelingMCP\\",
        "",
        "4. Re-run this script to register it:  python setup_pbi_mcp.py"
    )


# ─── 5. Main ──────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(
        description="Install & configure the Power BI Modeling MCP for Claude Code"
    )
    ap.add_argument("--check", action="store_true",
                    help="Check installation status only — no changes")
    ap.add_argument("--force", action="store_true",
                    help="Reinstall even if already configured")
    args = ap.parse_args()

    _banner("Power BI Modeling MCP  —  Setup")

    # ── Status snapshot ───────────────────────────────────────────────────────
    configured, cfg_exe = read_mcp_json_config()
    existing_exe        = find_installed_exe()

    _step("Current status")
    print(f"    .mcp.json configured  : {'YES — ' + str(cfg_exe) if configured else 'NO'}")
    print(f"    MCP exe on disk       : {existing_exe or 'NOT FOUND'}")

    if args.check:
        print()
        if configured and cfg_exe:
            _ok("MCP is installed and configured correctly")
        elif configured and not cfg_exe:
            _warn(".mcp.json has an entry but the exe path is invalid — re-run without --check")
        elif existing_exe and not configured:
            _warn("Exe found on disk but not registered in .mcp.json — re-run without --check")
        else:
            _warn("MCP not installed.  Run:  python setup_pbi_mcp.py")
        return 0

    # ── Already fully good ────────────────────────────────────────────────────
    if configured and cfg_exe and not args.force:
        print_already_configured(cfg_exe)
        return 0

    # ── Exe on disk but not registered ───────────────────────────────────────
    if existing_exe and not args.force:
        _step(f"Found existing installation — registering it")
        write_mcp_json(existing_exe)
        print_next_steps()
        return 0

    # ── Need to install ───────────────────────────────────────────────────────
    _step("MCP not found — downloading now")
    exe_path = download_and_extract_vsix()

    if exe_path is None:
        _fail("Could not install automatically")
        print()
        print_manual_steps()
        return 1

    # ── Register ──────────────────────────────────────────────────────────────
    _step("Registering in .mcp.json")
    write_mcp_json(exe_path)
    print_next_steps()
    return 0


if __name__ == "__main__":
    sys.exit(main())
