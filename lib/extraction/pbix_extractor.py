"""
PBIX Extraction Module

Extracts report pages, semantic model metadata, and static screenshots from
Power BI .pbix files (ZIP archives). Enables Claude to analyze PBIX files
without opening Power BI Desktop.

.pbix uses the same report JSON structure as .pbip but wraps it in a single
ZIP file. All visual.json parsing functions are reused from pbip_extractor
via direct import.

Model metadata: the DataModel entry in modern .pbix files is XPress9-compressed
binary (unreadable). Older files may include a DataModelSchema JSON file (BIM
format) alongside it — the extractor tries to parse it and falls back gracefully
to an empty model.

Static screenshots: every .pbix embeds pre-rendered PNG thumbnails under
Report/StaticResources/RegisteredResources/ — extracted without opening Desktop.
"""

import json
import re
import zipfile
from pathlib import Path
from typing import Optional

from lib.extraction.pbip_extractor import (
    _parse_visual_json,       # visual.json dict → structured visual info (same schema)
    build_dax_queries,        # (pages, model) → dax_queries list
    _capture_pbi_desktop_screenshots,  # UIAutomation if Desktop is open
)

# Roles whose queryRefs carry measures vs dimensions (same as pbip_extractor)
_MEASURE_ROLES = frozenset({
    'Y', 'Values', 'Tooltips', 'Size', 'Color', 'Details',
    'TargetValue', 'TrendLine', 'Indicator', 'Goal',
    'KPIIndicatorValue', 'KPIStatus', 'KPITrend',
})
_COLUMN_ROLES = frozenset({
    'Category', 'X', 'Axis', 'Legend', 'Series', 'Small multiples',
    'Rows', 'Columns', 'Group', 'Location',
})


# ---------------------------------------------------------------------------
# 1. Discover report pages
# ---------------------------------------------------------------------------

def _parse_queryref(query_ref: str) -> tuple:
    """
    Parse an old-format queryRef string into (entity, property, is_aggregation).

    Old-format projections use strings like:
      - "Table.MeasureName"           → entity="Table", prop="MeasureName", is_agg=False
      - "Sum(Table.Column Name)"      → entity="Table", prop="Column Name", is_agg=True
      - "Count(Table.Column)"         → entity="Table", prop="Column", is_agg=True
    """
    inner = query_ref.strip()
    is_agg = False
    agg_match = re.match(r'^\w+\((.+)\)$', inner)
    if agg_match:
        inner = agg_match.group(1)
        is_agg = True
    if '.' in inner:
        entity, prop = inner.split('.', 1)
        return entity.strip(), prop.strip(), is_agg
    return '', inner.strip(), is_agg


def _discover_pages_from_layout(zf: zipfile.ZipFile) -> list:
    """
    Parse pages from the monolithic Report/Layout JSON blob.

    Used for older PBIX files that do not have the split
    Report/definition/pages/ file structure.

    Layout.sections[] → pages; each section has visualContainers[] with
    a 'config' JSON string containing singleVisual data.
    """
    try:
        raw = zf.read('Report/Layout')
        data = json.loads(raw.decode('utf-16'))
    except Exception as e:
        print(f"  WARNING: Failed to parse Report/Layout: {e}")
        return []

    pages = []
    for section in data.get('sections', []):
        display_name = section.get('displayName') or section.get('name', 'Page')
        ordinal = section.get('ordinal', 999)
        is_hidden = section.get('visibility', 0) != 0

        visuals = []
        for vc in section.get('visualContainers', []):
            try:
                cfg = json.loads(vc.get('config', '{}'))
            except Exception:
                continue

            sv = cfg.get('singleVisual', {})
            vtype = sv.get('visualType', 'unknown') or 'unknown'

            # Extract title from objects.title (same path as PBIP visual.json)
            title = ''
            try:
                title_val = (sv.get('objects', {})
                             .get('title', [{}])[0]
                             .get('properties', {})
                             .get('text', {})
                             .get('expr', {})
                             .get('Literal', {})
                             .get('Value', ''))
                if title_val:
                    title = title_val.strip("'\"")
            except Exception:
                pass

            # Extract field bindings from old-format projections
            # projections: { role: [{"queryRef": "Entity.Property"}, ...] }
            measures = []
            columns = []
            seen: set = set()
            for role, items in sv.get('projections', {}).items():
                for item in (items if isinstance(items, list) else []):
                    qref = item.get('queryRef', '')
                    if not qref:
                        continue
                    entity, prop, is_agg = _parse_queryref(qref)
                    key = (entity, prop)
                    if key in seen or not prop:
                        continue
                    seen.add(key)
                    if role in _MEASURE_ROLES or is_agg:
                        measures.append({'name': prop, 'entity': entity})
                    else:
                        columns.append({'name': prop, 'entity': entity})

            # Approximate position from vc-level x/y/width/height
            width = vc.get('width', 0)
            height = vc.get('height', 0)

            visuals.append({
                'visual_id': cfg.get('name', ''),
                'visual_type': vtype,
                'title': title,
                'measures': measures,
                'columns': columns,
                'width': int(width),
                'height': int(height),
            })

        pages.append({
            'name': section.get('name', f'page_{ordinal}'),
            'display_name': display_name,
            'ordinal': ordinal,
            'is_hidden': is_hidden,
            'visuals': visuals,
        })

    # Sort by ordinal (already present from Layout)
    pages.sort(key=lambda p: p['ordinal'])
    return pages


def _discover_pages(zf: zipfile.ZipFile) -> list:
    """
    Discover report pages from the ZIP.

    Supports two PBIX formats:
      - Modern (PBIP-compatible): Report/definition/pages/<id>/page.json
      - Legacy (monolithic):      Report/Layout  (single UTF-16 JSON blob)

    Returns a list of page dicts in the same shape for both formats.
    """
    namelist = zf.namelist()

    # Detect format: if any page.json exists under definition/pages/, use split format
    has_definition_pages = any(
        re.match(r'^Report/definition/pages/[^/]+/page\.json$', e)
        for e in namelist
    )

    if has_definition_pages:
        return _discover_pages_split(zf, namelist)

    if 'Report/Layout' in namelist:
        print("  Detected legacy monolithic Report/Layout format")
        return _discover_pages_from_layout(zf)

    print("  WARNING: No recognisable page structure found in ZIP")
    return []


def _discover_pages_split(zf: zipfile.ZipFile, namelist: list) -> list:
    """
    Parse pages from the split Report/definition/pages/ structure (modern PBIX / PBIP-compat).
    Original implementation of _discover_pages.
    """
    page_json_re = re.compile(
        r'^Report/definition/pages/([^/]+)/page\.json$'
    )
    visual_json_re = re.compile(
        r'^Report/definition/pages/([^/]+)/visuals/([^/]+)/visual\.json$'
    )

    page_map: dict = {}
    for entry in namelist:
        m = page_json_re.match(entry)
        if not m:
            continue
        page_id = m.group(1)
        try:
            raw = zf.read(entry)
            page_cfg = json.loads(raw.decode('utf-8'))
        except Exception as e:
            print(f"  WARNING: Failed to parse {entry}: {e}")
            continue

        display_name = page_cfg.get('displayName') or page_cfg.get('name') or page_id
        ordinal = page_cfg.get('ordinal', 999)
        is_hidden = page_cfg.get('visibility', 0) != 0

        page_map[page_id] = {
            'name': page_id,
            'display_name': display_name,
            'ordinal': ordinal,
            'is_hidden': is_hidden,
            'visuals': [],
        }

    for entry in namelist:
        m = visual_json_re.match(entry)
        if not m:
            continue
        page_id = m.group(1)
        visual_id = m.group(2)
        if page_id not in page_map:
            continue
        try:
            raw = zf.read(entry)
            vis_cfg = json.loads(raw.decode('utf-8'))
            vis_info = _parse_visual_json(visual_id, vis_cfg)
            page_map[page_id]['visuals'].append(vis_info)
        except Exception as e:
            print(f"  WARNING: Failed to parse {entry}: {e}")

    return list(page_map.values())


# ---------------------------------------------------------------------------
# 2. Order pages
# ---------------------------------------------------------------------------

def _order_pages(zf: zipfile.ZipFile, pages: list) -> list:
    """
    Re-order pages using the pageOrder array in Report/definition/pages/pages.json
    (split/modern format only). Falls back to the ordinal field already set on each
    page dict — which is correct for both the split format and the legacy Layout format.
    """
    pages_json_entry = 'Report/definition/pages/pages.json'
    if pages_json_entry in zf.namelist():
        try:
            raw = zf.read(pages_json_entry)
            data = json.loads(raw.decode('utf-8'))
            page_order = data.get('pageOrder', [])
            if page_order:
                order_map = {name: i for i, name in enumerate(page_order)}
                pages = sorted(pages, key=lambda p: order_map.get(p['name'], 9999))
                for i, p in enumerate(pages):
                    p['ordinal'] = i
                return pages
        except Exception as e:
            print(f"  WARNING: Failed to parse pages.json order: {e}")

    # Fallback: ordinal from page.json (split) or sections[].ordinal (legacy Layout)
    pages.sort(key=lambda p: p['ordinal'])
    return pages


# ---------------------------------------------------------------------------
# 3. Parse model schema (BIM format — DataModelSchema entry)
# ---------------------------------------------------------------------------

def _parse_model_schema(zf: zipfile.ZipFile) -> dict:
    """
    Try to parse the DataModelSchema (BIM JSON) from the ZIP.

    Modern .pbix files store the model as XPress9-compressed binary in the
    DataModel entry (unreadable without the XPress9 library). Older files and
    some live-connected reports include a DataModelSchema JSON file (BIM
    format) alongside it.

    BIM structure:
        model.tables[].measures[].expression  → DAX expression
        model.relationships[]                 → table relationships

    Returns a model dict with tables/measures/relationships, or an empty model
    if DataModelSchema is absent or unparseable.
    """
    empty_model = {'tables': [], 'measures': [], 'relationships': []}
    namelist = zf.namelist()

    if 'DataModelSchema' not in namelist:
        print("  Model: DataModelSchema not present (modern binary format) "
              "— model metadata unavailable")
        return empty_model

    try:
        raw = zf.read('DataModelSchema')
        # DataModelSchema uses UTF-16 LE encoding (with BOM)
        text = raw.decode('utf-16')
        bim = json.loads(text)
    except Exception as e:
        print(f"  WARNING: Failed to parse DataModelSchema: {e}")
        return empty_model

    tables_out = []
    measures_out = []
    relationships_out = []

    model_node = bim.get('model', {})

    # Parse tables and their measures
    for table in model_node.get('tables', []):
        t_name = table.get('name', '')
        if not t_name:
            continue

        columns = []
        for col in table.get('columns', []):
            col_name = col.get('name', '')
            col_type = col.get('dataType', 'unknown')
            if col_name:
                columns.append({'name': col_name, 'dataType': col_type})

        tables_out.append({'name': t_name, 'columns': columns})

        # Measures are attached to tables in BIM format
        for measure in table.get('measures', []):
            m_name = measure.get('name', '')
            # BIM expression may be a plain string or a dict with a 'value' key
            m_expr = measure.get('expression', '')
            if isinstance(m_expr, dict):
                m_expr = m_expr.get('value', '')
            if m_name:
                measures_out.append({
                    'table': t_name,
                    'name': m_name,
                    'dax': str(m_expr).strip(),
                })

    # Parse relationships
    for rel in model_node.get('relationships', []):
        from_table = rel.get('fromTable', '')
        from_col = rel.get('fromColumn', '')
        to_table = rel.get('toTable', '')
        to_col = rel.get('toColumn', '')
        cardinality = rel.get('cardinality', '')
        if from_table and to_table:
            relationships_out.append({
                'from_table': from_table,
                'from_column': from_col,
                'to_table': to_table,
                'to_column': to_col,
                'cardinality': cardinality,
            })

    print(f"  Model: {len(tables_out)} tables, {len(measures_out)} measures, "
          f"{len(relationships_out)} relationships (from DataModelSchema)")

    return {
        'tables': tables_out,
        'measures': measures_out,
        'relationships': relationships_out,
    }


# ---------------------------------------------------------------------------
# 4. Static screenshot extraction
# ---------------------------------------------------------------------------

def _extract_static_screenshots(zf: zipfile.ZipFile, pages: list) -> dict:
    """
    Extract pre-rendered PNG thumbnails from Report/StaticResources/RegisteredResources/.

    Every .pbix embeds static resource PNGs saved at the last publish/save
    (~800-1200 px wide). These are matched to visible pages by name similarity
    and extracted to temp/pbix_page_N.png.

    Matching strategy:
      Pass 1: exact normalised name match
      Pass 2: word-overlap similarity (threshold 0.3)
      Pass 3: ordinal fallback — assign remaining PNGs by index order

    Returns dict: slide_number (1-based) -> "temp/pbix_page_N.png"
    or empty dict if no PNGs are found.
    """
    namelist = zf.namelist()

    # Collect PNG entries from RegisteredResources, excluding icons/logos:
    #   - Files with ".svg" in their stem (SVG icons saved as PNG, e.g. "icon.svg12345.png")
    #   - Square or portrait images (logos are square; page thumbnails are landscape 16:9)
    try:
        from PIL import Image as _PILImage
        import io as _io
        _pil_available = True
    except ImportError:
        _pil_available = False

    png_entries: list = []
    for entry in namelist:
        if not (entry.startswith('Report/StaticResources/RegisteredResources/')
                and entry.endswith('.png')):
            continue
        stem = Path(entry).stem
        # Skip SVG-derived icons (name contains ".svg")
        if '.svg' in stem.lower():
            continue
        # Skip square/portrait images (not page thumbnails) using PIL if available
        if _pil_available:
            try:
                raw = zf.read(entry)
                img = _PILImage.open(_io.BytesIO(raw))
                w, h = img.size
                # Page thumbnails are landscape; skip square (ratio ≤ 1.2) images
                if h == 0 or (w / h) <= 1.2:
                    continue
            except Exception:
                pass  # If we can't check, include it anyway
        png_entries.append((entry, stem))

    if not png_entries:
        print("  No static page-thumbnail PNGs found in ZIP")
        return {}

    print(f"  Found {len(png_entries)} static PNG(s) in ZIP")

    def _norm(s: str) -> str:
        """Lowercase, replace non-word chars with underscores."""
        s = re.sub(r'[^\w\s]', ' ', s)
        s = re.sub(r'[\s_]+', '_', s.lower().strip())
        return s.strip('_')

    def _word_overlap(page_norm: str, png_stem: str) -> float:
        """Word-overlap Jaccard score between normalised strings."""
        page_words = set(page_norm.split('_'))
        png_words = set(png_stem.split('_'))
        if not page_words or not png_words:
            return 0.0
        overlap = page_words & png_words
        return len(overlap) / max(len(page_words), len(png_words))

    Path('temp').mkdir(exist_ok=True)
    image_map: dict = {}
    used_entries: set = set()

    # Only match non-hidden pages
    visible_pages = [p for p in pages if not p.get('is_hidden', False)]

    # Pass 1: exact normalised name match
    for page_idx, page in enumerate(visible_pages):
        slide_num = page_idx + 1
        page_norm = _norm(page['display_name'])
        for entry, stem in png_entries:
            if _norm(stem) == page_norm and entry not in used_entries:
                out_path = f"temp/pbix_page_{slide_num}.png"
                with open(out_path, 'wb') as f:
                    f.write(zf.read(entry))
                image_map[slide_num] = out_path
                used_entries.add(entry)
                break

    # Pass 2: word-overlap for still-unmatched pages
    for page_idx, page in enumerate(visible_pages):
        slide_num = page_idx + 1
        if slide_num in image_map:
            continue
        page_norm = _norm(page['display_name'])
        remaining = [(e, s) for e, s in png_entries if e not in used_entries]
        if not remaining:
            break
        best_entry, best_stem = max(
            remaining,
            key=lambda es: _word_overlap(page_norm, _norm(es[1]))
        )
        score = _word_overlap(page_norm, _norm(best_stem))
        if score > 0.3:
            out_path = f"temp/pbix_page_{slide_num}.png"
            with open(out_path, 'wb') as f:
                f.write(zf.read(best_entry))
            image_map[slide_num] = out_path
            used_entries.add(best_entry)

    # Pass 3: ordinal fallback — assign remaining PNGs sequentially
    remaining_entries = [(e, s) for e, s in png_entries if e not in used_entries]
    for page_idx, page in enumerate(visible_pages):
        slide_num = page_idx + 1
        if slide_num in image_map:
            continue
        if not remaining_entries:
            break
        entry, stem = remaining_entries.pop(0)
        out_path = f"temp/pbix_page_{slide_num}.png"
        with open(out_path, 'wb') as f:
            f.write(zf.read(entry))
        image_map[slide_num] = out_path

    print(f"  Extracted {len(image_map)} of {len(visible_pages)} page screenshots")
    return image_map


# ---------------------------------------------------------------------------
# 5. Slide type classification (inlined — avoids import from main script)
# ---------------------------------------------------------------------------

def _classify_slide_type(title: str) -> str:
    """Keyword-based slide type classification (mirrors classify_slide_type in main script)."""
    name_lower = title.lower()
    if 'trend' in name_lower or 'over time' in name_lower:
        return 'trends'
    elif 'leaderboard' in name_lower or 'top' in name_lower:
        return 'leaderboard'
    elif 'health' in name_lower or 'overview' in name_lower:
        return 'health_check'
    elif 'habit' in name_lower or 'frequency' in name_lower:
        return 'habit_formation'
    elif 'license' in name_lower or 'priority' in name_lower:
        return 'license_priority'
    else:
        return 'general'


# ---------------------------------------------------------------------------
# 6. Main entry point
# ---------------------------------------------------------------------------

def prepare_pbix_for_claude_analysis(pbix_path: str) -> str:
    """
    Main entry point. Opens the .pbix ZIP and orchestrates all extraction steps.

    Steps:
      1. Open .pbix as zipfile.ZipFile
      2. _discover_pages(zf)           → pages list
      3. _order_pages(zf, pages)       → sorted pages (by pages.json pageOrder)
      4. _parse_model_schema(zf)       → model dict (tables/measures/relationships)
      5. _extract_static_screenshots(zf, pages) → dict {slide_num: path}
      6. If no screenshots: _capture_pbi_desktop_screenshots(pages, pbix.stem)
      7. build_dax_queries(pages, model) → dax_queries list
      8. Filter hidden pages; assign slide_numbers 1..N
      9. Write temp/analysis_request.json  (source_type: "pbix")
      10. Write temp/pbip_context.json     (same schema as pbip; key "pbix_path")

    Returns path to analysis_request.json.
    """
    print("=" * 70)
    print("PREPARING PBIX FILE FOR CLAUDE ANALYSIS")
    print("=" * 70)

    pbix_path = Path(pbix_path)
    if not pbix_path.exists():
        raise FileNotFoundError(f"PBIX file not found: {pbix_path}")
    if pbix_path.suffix.lower() != '.pbix':
        raise ValueError(f"Expected a .pbix file, got: {pbix_path}")

    print(f"\nFile: {pbix_path}")

    def _safe(text: str) -> str:
        """Strip non-ASCII-safe characters for Windows terminal output."""
        return re.sub(r'[^\x00-\x7F\x80-\xFF]', '?', str(text))

    try:
        zf = zipfile.ZipFile(str(pbix_path), 'r')
    except zipfile.BadZipFile as e:
        raise ValueError(f"File is not a valid .pbix (ZIP) archive: {e}")

    with zf:
        # Step 1 & 2: Discover and order report pages
        print("\nDiscovering report pages...")
        pages = _discover_pages(zf)
        pages = _order_pages(zf, pages)
        visible_pages = [p for p in pages if not p['is_hidden']]
        print(f"  Found {len(visible_pages)} visible pages "
              f"({len(pages) - len(visible_pages)} hidden)")
        for p in visible_pages:
            n_visuals = len(p['visuals'])
            print(f"  - {_safe(p['display_name'])} ({n_visuals} visuals)")

        # Step 3: Parse model schema
        print("\nParsing model schema...")
        model = _parse_model_schema(zf)

        # Step 4: Build DAX queries
        print("\nBuilding DAX queries for each page...")
        dax_queries = build_dax_queries(visible_pages, model)
        total_queries = sum(len(d['queries']) for d in dax_queries)
        print(f"  Generated {total_queries} queries across {len(dax_queries)} pages")

        # Step 5: Extract static screenshots from ZIP
        print("\nExtracting static screenshots from PBIX...")
        page_images = _extract_static_screenshots(zf, pages)

    # Step 6: If no static screenshots, try live Power BI Desktop capture
    if not page_images:
        print("\nNo static screenshots found — trying Power BI Desktop...")
        page_images = _capture_pbi_desktop_screenshots(visible_pages, pbix_path.stem)
        if page_images:
            print(f"  Captured {len(page_images)} screenshots from Power BI Desktop")
        else:
            print("  No images found — slides will be text/DAX-only")
            print("  Tip: Open the .pbix in Power BI Desktop, then re-run --prepare")

    # Step 7 & 8: Write output files
    Path('temp').mkdir(exist_ok=True)

    slides_meta = []
    for i, page in enumerate(visible_pages, start=1):
        slides_meta.append({
            'slide_number': i,
            'title': page['display_name'],
            'image_path': page_images.get(i),
            'slide_type': _classify_slide_type(page['display_name']),
            'source_type': 'pbix',
        })

    request_data = {
        'source_file': str(pbix_path),
        'source_type': 'pbix',
        'total_slides': len(slides_meta),
        'slides': slides_meta,
    }
    request_file = 'temp/analysis_request.json'
    with open(request_file, 'w', encoding='utf-8') as f:
        json.dump(request_data, f, indent=2)

    # pbip_context.json — same schema as PBIP; "pbix_path" key marks the source
    # trigger_claude_analysis() detects this file and routes to _trigger_pbip_analysis()
    context_data = {
        'pbix_path': str(pbix_path),
        'pages': visible_pages,
        'model': model,
        'dax_queries': dax_queries,
    }
    context_file = 'temp/pbip_context.json'
    with open(context_file, 'w', encoding='utf-8') as f:
        json.dump(context_data, f, indent=2)

    print(f"\nOK Analysis request saved to: {request_file}")
    print(f"OK PBIX context saved to:      {context_file}")

    return request_file
