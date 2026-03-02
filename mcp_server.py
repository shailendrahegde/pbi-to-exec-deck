#!/usr/bin/env python3
"""
MCP server for pbi-to-exec-deck.

Exposes three tools to Claude Code:
  - prepare_dashboard  : extract slides/pages from a PBI export → temp/
  - build_deck         : render temp/claude_insights.json → executive PPTX
  - check_pbi_connection : verify the Power BI Modeling MCP is reachable

Register in .mcp.json (run: python setup_pbi_mcp.py --register-deck-server):

  {
    "mcpServers": {
      "pbi-to-exec-deck": {
        "command": "python",
        "args": ["mcp_server.py"],
        "cwd": "<absolute path to this repo>"
      }
    }
  }
"""

import json
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Lazy imports of project modules so the server starts even if optional
# heavy deps (e.g. PyMuPDF) are not yet installed.
# ---------------------------------------------------------------------------

mcp = FastMCP("pbi-to-exec-deck")


# ---------------------------------------------------------------------------
# Tool: prepare_dashboard
# ---------------------------------------------------------------------------

@mcp.tool()
def prepare_dashboard(source_path: str) -> str:
    """
    Extract dashboard images and metadata from a Power BI export.

    Supports .pptx, .pdf, .pbip, and .pbix source files.
    Writes extracted images and analysis_request.json into the temp/ directory.

    Returns a JSON string with:
      - source_path   : echoed input path
      - source_type   : detected file type (pptx | pdf | pbip | pbix)
      - total_slides  : number of slides/pages extracted
      - slides        : list of {slide_number, title, image_path, slide_type}
      - pbip_context  : true if temp/pbip_context.json was also written
      - mcp_ready     : true if the Power BI Modeling MCP is configured (PBIP mode)
    """
    try:
        from convert_dashboard_claude import (
            prepare_for_claude_analysis,
            detect_file_type,
            _is_mcp_ready,
        )
    except ImportError as exc:
        return json.dumps({"error": f"Import failed: {exc}"})

    try:
        source_type = detect_file_type(source_path)
    except ValueError as exc:
        return json.dumps({"error": str(exc)})

    try:
        request_file = prepare_for_claude_analysis(source_path)
    except Exception as exc:
        return json.dumps({"error": f"Extraction failed: {exc}"})

    with open(request_file, "r", encoding="utf-8") as f:
        request = json.load(f)

    pbip_context = Path("temp/pbip_context.json").exists()
    mcp_ready = _is_mcp_ready() if source_type in ("pbip", "pbix") else False

    return json.dumps(
        {
            "source_path": source_path,
            "source_type": source_type,
            "total_slides": request.get("total_slides", 0),
            "slides": request.get("slides", []),
            "pbip_context": pbip_context,
            "mcp_ready": mcp_ready,
        },
        indent=2,
    )


# ---------------------------------------------------------------------------
# Tool: build_deck
# ---------------------------------------------------------------------------

@mcp.tool()
def build_deck(output_path: str = "") -> str:
    """
    Render temp/claude_insights.json into an executive PPTX presentation.

    Call this after you have written your insights to temp/claude_insights.json.

    Args:
        output_path: Where to save the .pptx file.
                     If omitted, defaults to <source>_executive.pptx
                     in the same directory as the original source file.

    Returns a JSON string with:
      - output_path   : absolute path of the created file
      - slide_count   : number of content slides rendered
      - deck_title    : title written on the cover slide
      - validation    : brief validation report from the constitution checker
    """
    try:
        from convert_dashboard_claude import (
            build_presentation_from_insights,
            generate_output_filename,
        )
    except ImportError as exc:
        return json.dumps({"error": f"Import failed: {exc}"})

    insights_file = "temp/claude_insights.json"
    if not Path(insights_file).exists():
        return json.dumps(
            {
                "error": (
                    f"{insights_file} not found. "
                    "Write your insights there before calling build_deck."
                )
            }
        )

    try:
        with open(insights_file, "r", encoding="utf-8") as f:
            insights_data = json.load(f)
    except json.JSONDecodeError as exc:
        return json.dumps({"error": f"Invalid JSON in {insights_file}: {exc}"})

    # Resolve source path from the analysis request written during prepare_dashboard
    request_file = "temp/analysis_request.json"
    if not Path(request_file).exists():
        return json.dumps(
            {
                "error": (
                    f"{request_file} not found. "
                    "Run prepare_dashboard before build_deck."
                )
            }
        )

    with open(request_file, "r", encoding="utf-8") as f:
        request = json.load(f)
    source_path = request["source_file"]

    resolved_output = output_path.strip() or generate_output_filename(source_path)

    try:
        # Capture validation output
        import io as _io
        import contextlib

        buf = _io.StringIO()
        with contextlib.redirect_stdout(buf):
            build_presentation_from_insights(source_path, resolved_output, insights_file)
        build_log = buf.getvalue()
    except Exception as exc:
        return json.dumps({"error": f"Build failed: {exc}"})

    # Extract a compact validation line from the log
    validation_lines = [
        line.strip()
        for line in build_log.splitlines()
        if line.strip() and not line.startswith("=") and not line.startswith("-")
    ]
    validation_summary = " | ".join(validation_lines[-6:]) if validation_lines else "OK"

    return json.dumps(
        {
            "output_path": str(Path(resolved_output).resolve()),
            "slide_count": len(insights_data.get("slides", [])),
            "deck_title": insights_data.get("deck_title", ""),
            "validation": validation_summary,
        },
        indent=2,
    )


# ---------------------------------------------------------------------------
# Tool: check_pbi_connection
# ---------------------------------------------------------------------------

@mcp.tool()
def check_pbi_connection() -> str:
    """
    Check whether the Power BI Modeling MCP server is installed and reachable.

    Returns a JSON string with:
      - configured  : true if .mcp.json references the powerbi-modeling server
      - exe_valid   : true if the registered executable path exists on disk
      - ready       : true only when both configured AND exe_valid are true
      - message     : human-readable status or instructions
    """
    mcp_json = Path(".mcp.json")
    configured = False
    exe_valid = False
    exe_path = ""

    if mcp_json.exists():
        try:
            cfg = json.loads(mcp_json.read_text(encoding="utf-8"))
            server = cfg.get("mcpServers", {}).get("powerbi-modeling")
            if server:
                configured = True
                exe_path = server.get("command", "")
                exe_valid = Path(exe_path).exists()
        except Exception:
            pass

    if configured and exe_valid:
        message = "Power BI Modeling MCP is ready — DAX query mode enabled."
    elif configured and not exe_valid:
        message = (
            f"MCP is registered in .mcp.json but the executable was not found "
            f"at '{exe_path}'. Re-run: python setup_pbi_mcp.py --force"
        )
    else:
        message = (
            "Power BI Modeling MCP is not installed. "
            "Run: python setup_pbi_mcp.py  then restart Claude Code."
        )

    return json.dumps(
        {
            "configured": configured,
            "exe_valid": exe_valid,
            "ready": configured and exe_valid,
            "message": message,
        },
        indent=2,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run(transport="stdio")
