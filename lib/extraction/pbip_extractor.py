"""
PBIP Extraction Module

Extracts report pages, semantic model metadata, and builds DAX queries
for Power BI PBIP projects. Enables the assistant to query the live in-memory
model via the powerbi-modeling MCP instead of reading screenshots.
"""

import json
import re
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# TMDL Parser
# ---------------------------------------------------------------------------

def _parse_tmdl_file(filepath: Path) -> dict:
    """
    Parse a single .tmdl file and return extracted tables/measures/relationships.

    TMDL is an indentation-based plain-text format.  A typical table file:

        table Sales
            column OrderId
                dataType: int64

            measure 'Total Revenue' =
                    SUM(Sales[Amount])
                formatString: #,##0.00

    We scan line-by-line, tracking the current table context and accumulating
    multi-line measure DAX expressions.
    """
    result = {
        'tables': [],      # {name, columns: [{name, dataType}], measures: [{name, dax}]}
        'relationships': [],  # {from_table, from_column, to_table, to_column, cardinality}
    }

    try:
        text = filepath.read_text(encoding='utf-8', errors='replace')
    except Exception:
        return result

    lines = text.splitlines()

    current_table: Optional[dict] = None
    current_column: Optional[dict] = None
    current_measure: Optional[dict] = None
    in_measure_expr = False  # True while accumulating a multi-line expression
    measure_expr_lines: list = []

    def flush_measure():
        """Finalise any in-progress measure and attach it to current_table."""
        nonlocal current_measure, in_measure_expr, measure_expr_lines
        if current_measure and current_table is not None:
            if measure_expr_lines:
                current_measure['dax'] = '\n'.join(measure_expr_lines).strip()
            current_table['measures'].append(current_measure)
        current_measure = None
        in_measure_expr = False
        measure_expr_lines = []

    def flush_column():
        nonlocal current_column
        if current_column and current_table is not None:
            current_table['columns'].append(current_column)
        current_column = None

    def get_indent(line: str) -> int:
        return len(line) - len(line.lstrip('\t '))

    in_relationship = False
    rel_buf: dict = {}

    for raw_line in lines:
        stripped = raw_line.strip()
        if not stripped:
            continue

        indent = get_indent(raw_line)

        # ---- Relationship block detection ------------------------------------
        if stripped.startswith('relationship'):
            in_relationship = True
            rel_buf = {}
            flush_measure()
            flush_column()
            continue

        if in_relationship:
            if stripped.startswith('fromTable:'):
                rel_buf['from_table'] = stripped.split(':', 1)[1].strip().strip("'")
            elif stripped.startswith('fromColumn:'):
                rel_buf['from_column'] = stripped.split(':', 1)[1].strip().strip("'")
            elif stripped.startswith('toTable:'):
                rel_buf['to_table'] = stripped.split(':', 1)[1].strip().strip("'")
            elif stripped.startswith('toColumn:'):
                rel_buf['to_column'] = stripped.split(':', 1)[1].strip().strip("'")
            elif stripped.startswith('cardinality:'):
                rel_buf['cardinality'] = stripped.split(':', 1)[1].strip()
            elif stripped.startswith('crossFilteringBehavior:'):
                rel_buf['cross_filter'] = stripped.split(':', 1)[1].strip()
            # End of relationship block when we hit something at top indent
            elif indent == 0 and not stripped.startswith(('\t', ' ')):
                if rel_buf.get('from_table') and rel_buf.get('to_table'):
                    result['relationships'].append(rel_buf)
                in_relationship = False
                rel_buf = {}
            continue

        # ---- Table block detection -------------------------------------------
        table_match = re.match(r"^table\s+'?(.+?)'?\s*$", stripped)
        if table_match and indent == 0:
            # Finalise previous table
            flush_measure()
            flush_column()
            if current_table:
                result['tables'].append(current_table)
            current_table = {
                'name': table_match.group(1),
                'columns': [],
                'measures': [],
            }
            current_column = None
            current_measure = None
            continue

        if current_table is None:
            continue  # Skip lines outside any table context

        # ---- Measure definition ----------------------------------------------
        # Measure lines look like: measure 'Name' = <expr>  (or just = on next line)
        measure_match = re.match(r"^\s+measure\s+'?(.+?)'?\s*=\s*(.*)", raw_line)
        if measure_match:
            flush_measure()
            flush_column()
            m_name = measure_match.group(1).strip().strip("'")
            m_expr_start = measure_match.group(2).strip()
            current_measure = {'name': m_name, 'dax': ''}
            in_measure_expr = True
            if m_expr_start:
                measure_expr_lines = [m_expr_start]
            else:
                measure_expr_lines = []
            continue

        # ---- Accumulate measure expression lines ----------------------------
        if in_measure_expr and current_measure:
            # A property like "formatString:", "displayFolder:", "annotation" ends the expr
            prop_match = re.match(r"^\s+(formatString|displayFolder|annotation|isHidden|"
                                  r"description|kpiStatusExpression|kpiTargetExpression"
                                  r"|lineageTag|summarizeBy|dataCategory)\s*:", raw_line)
            if prop_match:
                flush_measure()
            else:
                # Continuation of DAX expression
                measure_expr_lines.append(stripped)
            continue

        # ---- Column definition -----------------------------------------------
        col_match = re.match(r"^\s+column\s+'?(.+?)'?\s*$", raw_line)
        if col_match:
            flush_column()
            flush_measure()
            current_column = {'name': col_match.group(1).strip().strip("'"), 'dataType': 'unknown'}
            continue

        if current_column:
            if stripped.startswith('dataType:'):
                current_column['dataType'] = stripped.split(':', 1)[1].strip()

    # ---- Flush last items ---------------------------------------------------
    flush_measure()
    flush_column()
    if current_table:
        result['tables'].append(current_table)
    if in_relationship and rel_buf.get('from_table'):
        result['relationships'].append(rel_buf)

    return result


# ---------------------------------------------------------------------------
# 1. Discover report pages
# ---------------------------------------------------------------------------

def discover_report_pages(pbip_root: str) -> list:
    """
    Walk the Report definition directory and return a list of pages with their
    visuals and field bindings.

    Each returned dict has:
        {
          "name": str,           # system page name (folder name)
          "display_name": str,   # human-readable name from page.json
          "ordinal": int,        # display order
          "is_hidden": bool,
          "visuals": [
            {
              "visual_id": str,
              "visual_type": str,   # barChart, card, table, matrix, …
              "title": str,
              "measures": [str],    # measure names referenced by this visual
              "columns": [str],     # column names referenced by this visual
              "width": int,
              "height": int,
            }
          ]
        }
    """
    root = Path(pbip_root)

    # Derive expected .Report folder name from the .pbip filename
    # e.g. "LIGHT AI-in-One.pbip" → "LIGHT AI-in-One.Report"
    expected_report_name = None
    if root.suffix.lower() == '.pbip':
        expected_report_name = root.stem + '.Report'

    # Find the .Report folder
    report_dir = None

    def _find_report_dir(search_dir: Path) -> Path | None:
        """Search a directory for the matching .Report folder."""
        if not search_dir.is_dir():
            return None
        # Prefer exact name match when we know the expected folder name
        if expected_report_name:
            exact = search_dir / expected_report_name
            if exact.is_dir():
                return exact
        # Fallback: first .Report folder found
        for candidate in search_dir.iterdir():
            if candidate.is_dir() and candidate.name.endswith('.Report'):
                return candidate
        return None

    # 1. Search inside root (if root is a directory)
    if root.is_dir():
        report_dir = _find_report_dir(root)

    # 2. If root itself is the .Report folder
    if report_dir is None and root.name.endswith('.Report') and root.is_dir():
        report_dir = root

    # 3. Search sibling directories (parent of root)
    if report_dir is None:
        report_dir = _find_report_dir(root.parent)

    if report_dir is None:
        print(f"  WARNING: Could not find .Report directory under '{pbip_root}'")
        return []

    pages_dir = report_dir / 'definition' / 'pages'
    if not pages_dir.exists():
        print(f"  WARNING: Pages directory not found: {pages_dir}")
        return []

    pages = []

    for page_dir in sorted(pages_dir.iterdir()):
        if not page_dir.is_dir():
            continue

        page_json_path = page_dir / 'page.json'
        if not page_json_path.exists():
            continue

        try:
            page_cfg = json.loads(page_json_path.read_text(encoding='utf-8'))
        except Exception as e:
            print(f"  WARNING: Failed to parse {page_json_path}: {e}")
            continue

        display_name = page_cfg.get('displayName') or page_cfg.get('name') or page_dir.name
        ordinal = page_cfg.get('ordinal', 999)
        is_hidden = page_cfg.get('visibility', 0) != 0  # 0 = visible, 1+ = hidden

        visuals = []
        visuals_dir = page_dir / 'visuals'
        if visuals_dir.exists():
            for visual_dir in sorted(visuals_dir.iterdir()):
                if not visual_dir.is_dir():
                    continue
                visual_json_path = visual_dir / 'visual.json'
                if not visual_json_path.exists():
                    continue

                try:
                    vis_cfg = json.loads(visual_json_path.read_text(encoding='utf-8'))
                    vis_info = _parse_visual_json(visual_dir.name, vis_cfg)
                    visuals.append(vis_info)
                except Exception as e:
                    print(f"  WARNING: Failed to parse {visual_json_path}: {e}")

        pages.append({
            'name': page_dir.name,
            'display_name': display_name,
            'ordinal': ordinal,
            'is_hidden': is_hidden,
            'page_width': page_cfg.get('width'),
            'page_height': page_cfg.get('height'),
            'visuals': visuals,
        })

    # Determine correct page order.
    # Priority 1: pages/pages.json pageOrder array — the authoritative source.
    #   This file is part of the PBIP v2 format and lists page folder names in
    #   the exact order Power BI Desktop displays them in the tab bar.
    # Priority 2: RegisteredResources PNG list in report.json — used as a
    #   fallback when pages.json is absent (older PBIP versions).
    # Priority 3: ordinal field in page.json (rarely set in PBIP v2).
    pages_json_path = pages_dir / 'pages.json'
    if pages_json_path.exists():
        pages = _order_pages_by_pages_json(pages, pages_json_path)
    else:
        report_json_path = report_dir / 'definition' / 'report.json'
        if report_json_path.exists():
            pages = _order_pages_by_resource_packages(pages, report_json_path)
        else:
            pages.sort(key=lambda p: p['ordinal'])
    return pages


def _order_pages_by_pages_json(pages: list, pages_json_path: Path) -> list:
    """
    Re-order pages using the `pageOrder` array in pages/pages.json.

    This is the authoritative source for PBIP v2: the array lists page folder
    names in the exact sequence Power BI Desktop displays them in the tab bar.
    Pages not present in pageOrder (should not happen in valid PBIP) are appended.
    """
    try:
        data = json.loads(pages_json_path.read_text(encoding='utf-8'))
    except Exception:
        return pages

    page_order = data.get('pageOrder', [])
    if not page_order:
        return pages

    order_map = {name: i for i, name in enumerate(page_order)}
    # Sort by position in pageOrder; pages absent from the list go to the end.
    sorted_pages = sorted(pages, key=lambda p: order_map.get(p['name'], 9999))

    # Stamp ordinals so the rest of the pipeline sees them
    for i, p in enumerate(sorted_pages):
        p['ordinal'] = i

    return sorted_pages


def _order_pages_by_resource_packages(pages: list, report_json_path: Path) -> list:
    """
    Re-order pages to match Power BI Desktop's tab bar order.

    In PBIP v2, page ordinals are often absent.  Power BI Desktop then orders
    pages according to the RegisteredResources PNG list in report.json (the
    order these static helper images were registered mirrors the order pages
    were created / arranged in the report).

    Matching strategy:
      1. Normalise both the page display name and each PNG stem to lowercase
         underscore_words.
      2. Direct (exact) match first.
      3. Word-overlap with partial prefix scoring for the remaining pages
         (handles cases like "chat_web_usage_trends" -> "copilot_chat_usage_trends").
      4. Pages with no PNG match (e.g. a Glossary page) are appended at the end.
    """
    try:
        report_data = json.loads(report_json_path.read_text(encoding='utf-8'))
    except Exception:
        return pages

    # Extract page PNG stems from RegisteredResources in order.
    # Skip the icon PNG (contains ".svg" in its stem — icon saved as PNG).
    png_order: list = []
    for pkg in report_data.get('resourcePackages', []):
        if pkg.get('type') != 'RegisteredResources':
            continue
        for item in pkg.get('items', []):
            if item.get('type') == 'Image':
                name = item.get('name', '')
                if name.endswith('.png') and '.svg' not in name:
                    png_order.append(name[:-4])  # strip .png

    if not png_order:
        return pages

    def _norm(s: str) -> str:
        """Lowercase, remove non-word/space chars, collapse to underscores."""
        s = re.sub(r'[^\w\s]', ' ', s)
        s = re.sub(r'[\s_]+', '_', s.lower().strip())
        return s.strip('_')

    def _word_sim(w1: str, w2: str) -> float:
        """Longest common prefix fraction (handles s/z spelling variants)."""
        if w1 == w2:
            return 1.0
        n = min(len(w1), len(w2))
        common = 0
        for i in range(n):
            if w1[i] == w2[i]:
                common += 1
            else:
                break
        return common / max(len(w1), len(w2))

    def _overlap_score(page_norm: str, png_stem: str) -> float:
        """Sum of best word-similarity scores between page words and PNG words."""
        page_words = page_norm.split('_')
        png_words  = png_stem.split('_')
        total = 0.0
        for pw in page_words:
            best = max((_word_sim(pw, sw) for sw in png_words), default=0.0)
            total += best
        return total

    # Build display-name → page lookup
    page_by_norm = {_norm(p['display_name']): p for p in pages}

    # Pass 1: direct (exact) matches
    result: list = [None] * len(png_order)
    matched_norms: set = set()
    for idx, stem in enumerate(png_order):
        if stem in page_by_norm:
            result[idx] = page_by_norm[stem]
            matched_norms.add(stem)

    # Pass 2: word-overlap for remaining PNG slots
    unmatched = [p for p in pages if _norm(p['display_name']) not in matched_norms]
    for idx, slot in enumerate(result):
        if slot is not None or not unmatched:
            continue
        stem = png_order[idx]
        best_page  = max(unmatched,
                         key=lambda p: _overlap_score(_norm(p['display_name']), stem))
        result[idx] = best_page
        unmatched.remove(best_page)
        matched_norms.add(_norm(best_page['display_name']))

    # Collect ordered pages (drop unfilled slots) and append unmatched remainder
    ordered = [p for p in result if p is not None]
    ordered.extend(unmatched)  # pages with no PNG match (e.g. Glossary) go last

    # Stamp ordinals so the rest of the pipeline sees them
    for i, p in enumerate(ordered):
        p['ordinal'] = i

    return ordered


def _parse_visual_json(visual_id: str, vis_cfg: dict) -> dict:
    """
    Extract structured info from a visual.json dict.

    Actual PBIP schema (as of 2024/2025 format):
      {
        "name": "...",
        "position": { "x", "y", "z", "height", "width", "tabOrder" },
        "visual": {
          "visualType": "clusteredColumnChart",
          "query": {
            "queryState": {
              "Y": { "projections": [ { "field": { "Measure": { "Expression": { "SourceRef": { "Entity": "TableName" } }, "Property": "MeasureName" } }, ... } ] },
              "Category": { "projections": [ { "field": { "Column": { ... } }, ... } ] }
            }
          },
          "objects": { ... }
        }
      }
    """
    pos = vis_cfg.get('position', {})
    info = {
        'visual_id': visual_id,
        'visual_type': 'unknown',
        'title': '',
        # measures: list of {name: str, entity: str}
        'measures': [],
        # columns: list of {name: str, entity: str}
        'columns': [],
        'width': pos.get('width', vis_cfg.get('width', 0)),
        'height': pos.get('height', vis_cfg.get('height', 0)),
    }

    visual_node = vis_cfg.get('visual', {})
    info['visual_type'] = visual_node.get('visualType', 'unknown') or 'unknown'

    # Extract title from visual.objects.title[0].properties.text
    try:
        title_val = (visual_node.get('objects', {})
                     .get('title', [{}])[0]
                     .get('properties', {})
                     .get('text', {})
                     .get('expr', {})
                     .get('Literal', {})
                     .get('Value', ''))
        if title_val:
            info['title'] = title_val.strip("'\"")
    except Exception:
        pass

    # Extract field projections from visual.query.queryState
    try:
        query_state = visual_node.get('query', {}).get('queryState', {})
        _extract_field_refs_v2(query_state, info)
    except Exception:
        pass

    return info


# Role names that carry measures (values/metrics) vs dimensions (axes/groupings)
_MEASURE_ROLES = frozenset({
    'Y', 'Values', 'Tooltips', 'Size', 'Color', 'Details',
    'TargetValue', 'TrendLine', 'Indicator', 'Goal',
    'KPIIndicatorValue', 'KPIStatus', 'KPITrend',
})
_COLUMN_ROLES = frozenset({
    'Category', 'X', 'Axis', 'Legend', 'Series', 'Small multiples',
    'Rows', 'Columns', 'Group', 'Location',
})


def _extract_field_refs_v2(query_state: dict, info: dict):
    """
    Parse visual.query.queryState to extract measure and column field refs.

    Each role in queryState has a 'projections' list. Each projection has a
    'field' dict that is one of:
      - { "Measure": { "Expression": { "SourceRef": { "Entity": "..." } }, "Property": "..." } }
      - { "Column": { "Expression": { "SourceRef": { "Entity": "..." } }, "Property": "..." } }
      - { "Aggregation": { ... } }   (implicit aggregation — treat column as group-by)
      - { "HierarchyLevel": { ... } } (date hierarchy — extract entity+level)
    """
    seen = set()  # deduplicate by (entity, property)

    for role, role_data in query_state.items():
        if not isinstance(role_data, dict):
            continue
        projections = role_data.get('projections', [])
        for proj in projections:
            field = proj.get('field', {})
            if not field:
                continue

            # Measure binding
            if 'Measure' in field:
                m = field['Measure']
                entity = m.get('Expression', {}).get('SourceRef', {}).get('Entity', '')
                prop = m.get('Property', '')
                key = (entity, prop)
                if entity and prop and key not in seen:
                    seen.add(key)
                    info['measures'].append({'name': prop, 'entity': entity})

            # Column binding
            elif 'Column' in field:
                c = field['Column']
                entity = c.get('Expression', {}).get('SourceRef', {}).get('Entity', '')
                prop = c.get('Property', '')
                key = (entity, prop)
                if entity and prop and key not in seen:
                    seen.add(key)
                    info['columns'].append({'name': prop, 'entity': entity})

            # Aggregation (e.g. COUNT of a column) — treat as column grouping
            elif 'Aggregation' in field:
                agg = field['Aggregation']
                inner_col = agg.get('Expression', {}).get('Column', {})
                entity = inner_col.get('Expression', {}).get('SourceRef', {}).get('Entity', '')
                prop = inner_col.get('Property', '')
                key = (entity, prop)
                if entity and prop and key not in seen:
                    seen.add(key)
                    info['columns'].append({'name': prop, 'entity': entity})

            # HierarchyLevel (date drill-down) — extract entity + level as column
            elif 'HierarchyLevel' in field:
                hl = field['HierarchyLevel']
                hier = hl.get('Expression', {}).get('Hierarchy', {})
                pv = hier.get('Expression', {}).get('PropertyVariationSource', {})
                entity = pv.get('Expression', {}).get('SourceRef', {}).get('Entity', '')
                prop = pv.get('Property', '')
                level = hl.get('Level', '')
                display = f"{prop} ({level})" if level else prop
                key = (entity, display)
                if entity and display and key not in seen:
                    seen.add(key)
                    info['columns'].append({'name': display, 'entity': entity})


# ---------------------------------------------------------------------------
# 2. Extract model metadata
# ---------------------------------------------------------------------------

def extract_model_metadata(pbip_root: str) -> dict:
    """
    Find the SemanticModel directory and parse all .tmdl files.

    Returns:
        {
          "tables": [...],
          "measures": [...],   # flattened list: {table, name, dax}
          "relationships": [...],
        }
    """
    root = Path(pbip_root)

    # Find the .SemanticModel folder
    # When passed a .pbip file, derive the expected folder name to avoid
    # picking up the wrong model when multiple PBIPs share a directory.
    model_dir = None
    expected_model_name = None
    if root.suffix.lower() == '.pbip':
        expected_model_name = root.stem + '.SemanticModel'
        search_root = root.parent
    else:
        search_root = root if root.is_dir() else root.parent

    for candidate in search_root.iterdir():
        if candidate.is_dir() and candidate.name.endswith('.SemanticModel'):
            if expected_model_name and candidate.name == expected_model_name:
                model_dir = candidate
                break
            elif not expected_model_name:
                model_dir = candidate
                break

    # Fallback: accept any .SemanticModel if exact match not found
    if model_dir is None:
        for candidate in search_root.iterdir():
            if candidate.is_dir() and candidate.name.endswith('.SemanticModel'):
                model_dir = candidate
                break

    if model_dir is None:
        print(f"  WARNING: Could not find .SemanticModel directory under '{pbip_root}'")
        return {'tables': [], 'measures': [], 'relationships': []}

    def_dir = model_dir / 'definition'
    if not def_dir.exists():
        def_dir = model_dir  # Some versions keep TMDL files directly in root

    all_tables = []
    all_relationships = []
    all_measures = []

    # Parse all .tmdl files recursively
    for tmdl_path in def_dir.rglob('*.tmdl'):
        parsed = _parse_tmdl_file(tmdl_path)
        for table in parsed['tables']:
            all_tables.append({
                'name': table['name'],
                'columns': table['columns'],
            })
            for measure in table['measures']:
                all_measures.append({
                    'table': table['name'],
                    'name': measure['name'],
                    'dax': measure['dax'],
                })
        all_relationships.extend(parsed['relationships'])

    return {
        'tables': all_tables,
        'measures': all_measures,
        'relationships': all_relationships,
    }


# ---------------------------------------------------------------------------
# 3. Build DAX queries
# ---------------------------------------------------------------------------

def build_dax_queries(pages: list, model: dict) -> list:
    """
    For each page, identify the measures referenced and generate targeted DAX
    queries. Returns a list of per-page query groups.

    Each entry:
        {
          "slide_number": int,
          "page_name": str,
          "queries": [
            {
              "label": str,
              "visual_type": str,
              "dax": str,             # ready-to-execute DAX query
              "measure_dax": str,     # underlying DAX expression (for context)
            }
          ]
        }

    measures in visuals are now dicts: {name: str, entity: str}
    columns in visuals are now dicts: {name: str, entity: str}
    """
    # Build a lookup: (table, measure_name) -> DAX expression
    measure_dax_lookup: dict = {}
    for m in model.get('measures', []):
        measure_dax_lookup[(m['table'], m['name'])] = m['dax']
        measure_dax_lookup[m['name']] = m['dax']  # name-only fallback

    result = []
    for slide_num, page in enumerate(pages, start=1):
        queries = []
        seen_queries: set = set()

        for visual in page.get('visuals', []):
            vtype = visual.get('visual_type', 'unknown')
            v_measures = visual.get('measures', [])  # list of {name, entity}
            v_columns = visual.get('columns', [])    # list of {name, entity}

            if vtype in ('slicer', 'actionButton', 'image', 'textbox', 'shape', 'unknown'):
                continue

            # Generate one query per measure binding (cap at 3 per visual)
            for m_ref in v_measures[:3]:
                m_name = m_ref['name'] if isinstance(m_ref, dict) else m_ref
                m_entity = m_ref.get('entity', '') if isinstance(m_ref, dict) else ''

                query_key = f"{vtype}:{m_entity}:{m_name}"
                if query_key in seen_queries:
                    continue
                seen_queries.add(query_key)

                # Look up the DAX expression for this measure
                dax_expr = (measure_dax_lookup.get((m_entity, m_name))
                            or measure_dax_lookup.get(m_name)
                            or '')

                dax_query = _build_dax_for_visual(vtype, m_name, m_entity, v_columns)

                if dax_query:
                    queries.append({
                        'label': f"{vtype} — {m_name}",
                        'visual_type': vtype,
                        'measure_name': m_name,
                        'measure_entity': m_entity,
                        'dax': dax_query,
                        'measure_dax': dax_expr or '(expression not found in model)',
                    })

            # Column-only visuals (table with no measures) — emit a simple query
            if not v_measures and v_columns and vtype in ('table', 'matrix', 'tableEx'):
                for col_ref in v_columns[:2]:
                    col_name = col_ref['name'] if isinstance(col_ref, dict) else col_ref
                    col_entity = col_ref.get('entity', '') if isinstance(col_ref, dict) else ''
                    query_key = f"{vtype}:col:{col_entity}:{col_name}"
                    if query_key in seen_queries or not col_entity:
                        continue
                    seen_queries.add(query_key)
                    dax = (f"EVALUATE\n"
                           f"TOPN(20, SUMMARIZECOLUMNS('{col_entity}'[{col_name}]))")
                    queries.append({
                        'label': f"table — {col_name}",
                        'visual_type': vtype,
                        'measure_name': None,
                        'measure_entity': col_entity,
                        'dax': dax,
                        'measure_dax': '',
                    })

        result.append({
            'slide_number': slide_num,
            'page_name': page['display_name'],
            'queries': queries,
        })

    return result


def _build_dax_for_visual(vtype: str, measure_name: str,
                           measure_entity: str, bound_columns: list) -> str:
    """
    Generate a DAX query appropriate for the given visual type.

    - card / KPI        → EVALUATE ROW(...)
    - chart / table     → EVALUATE TOPN(20, SUMMARIZECOLUMNS(group_col, measure))
    """
    # Quote entity names that contain spaces
    def quote(name: str) -> str:
        return f"'{name}'" if ' ' in name else name

    safe_measure = f"[{measure_name}]"

    if vtype in ('card', 'kpiVisual', 'singleValue', 'gauge', 'multiRowCard'):
        return f'EVALUATE\nROW("{measure_name}", {safe_measure})'

    # For charts and tables, try to find a grouping column
    group_col_dax = None
    for col_ref in bound_columns:
        col_name = col_ref['name'] if isinstance(col_ref, dict) else col_ref
        col_entity = col_ref.get('entity', '') if isinstance(col_ref, dict) else ''
        if col_entity:
            group_col_dax = f"{quote(col_entity)}[{col_name}]"
            break

    if group_col_dax:
        return (f'EVALUATE\n'
                f'TOPN(\n'
                f'    20,\n'
                f'    SUMMARIZECOLUMNS(\n'
                f'        {group_col_dax},\n'
                f'        "{measure_name}", {safe_measure}\n'
                f'    ),\n'
                f'    {safe_measure}, DESC\n'
                f')')
    else:
        return f'EVALUATE\nROW("{measure_name}", {safe_measure})'


# ---------------------------------------------------------------------------
# 4. Static resource image extraction (PBIP built-in)
# ---------------------------------------------------------------------------

def _capture_pbi_desktop_screenshots(pages: list, pbip_stem: str = '') -> dict:
    """
    Capture a screenshot of each report page from the running Power BI Desktop.

    Uses win32gui to locate the PBI Desktop window, then navigates between report
    pages by clicking directly on each page tab at the bottom of the window.
    Screenshots are captured with PIL.ImageGrab.

    Navigation approach: click each page tab at its approximate horizontal centre
    rather than using keyboard shortcuts (which require focus to be off all
    canvas visuals — unreliable).  Tab positions are computed assuming equal tab
    widths across the available tab-bar width.

    When multiple Power BI Desktop windows are open, the window whose title
    contains ``pbip_stem`` (the PBIP filename without extension) is preferred.
    PBI Desktop titles follow the pattern "<Report Name> - Power BI Desktop".

    Canvas bounds are estimated from the maximised window size:
      - Top   : ~130 px  (title bar + menu + quick-access toolbar + ribbon)
      - Bottom:  ~52 px  (page-tab strip + status bar)
      - Left  :  ~48 px  (left navigation rail, icon mode)
      - Right : auto-detected (crops out Visualizations/Fields/Format panels)

    Returns dict: slide_number (1-based) -> screenshot path string,
    or empty dict if Power BI Desktop is not found / capture fails.
    """
    try:
        import ctypes
        import time
        import win32gui
        import win32con
        import win32api
        from PIL import Image, ImageGrab
    except ImportError as e:
        print(f"  WARNING: Screenshot capture unavailable — missing dependency: {e}")
        return {}

    # --- 1. Find the correct Power BI Desktop window ---
    # PBI Desktop title formats vary by version:
    #   - "<Report Name> - Power BI Desktop"   (older builds)
    #   - "<Report Name>"                        (newer builds — no suffix)
    # Strategy: prefer a window whose title contains the PBIP stem; fall back
    # to any window with "Power BI Desktop" in the title; otherwise error.
    all_windows: list = []     # (hwnd, title) for every visible window with a title
    pbi_all: list = []         # subset with "Power BI Desktop" in title
    stem_matches: list = []    # subset whose title contains the PBIP stem

    def _enum_cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            return
        if not title:
            return
        all_windows.append((hwnd, title))
        if 'Power BI Desktop' in title:
            pbi_all.append((hwnd, title))

    win32gui.EnumWindows(_enum_cb, None)

    # Non-PBI Office app suffixes to exclude from stem matches
    _NON_PBI_SUFFIXES = ('- powerpoint', '- word', '- excel', '- outlook',
                         '- notepad', '- visual studio', '- code')

    # Build a normalised match key from the PBIP stem for fuzzy title matching
    if pbip_stem:
        match_key = re.sub(r'[^\w ]', '', pbip_stem.lower())[:25].strip()
        for hwnd, title in all_windows:
            norm = re.sub(r'[^\w ]', '', title.lower())
            if match_key and match_key in norm:
                # Exclude windows that are clearly a different Office application
                title_lower = title.lower()
                if not any(title_lower.endswith(s) or (s + ' ') in title_lower
                           for s in _NON_PBI_SUFFIXES):
                    stem_matches.append((hwnd, title))

    # Pick the best candidate
    # Priority order:
    #   1. Window whose title contains both the PBIP stem AND "Power BI Desktop"
    #   2. Any "Power BI Desktop" window (fallback when PBI omits the suffix)
    #   3. Stem match that isn't a known non-PBI application (last resort)
    pbi_hwnd = None
    stem_and_pbi = [hw for hw, t in stem_matches if 'Power BI Desktop' in t]
    if stem_and_pbi:
        pbi_hwnd = stem_and_pbi[0]
    elif stem_matches:
        # Stem found but "Power BI Desktop" not in title — newer PBI build
        pbi_hwnd = stem_matches[0][0]
    elif pbi_all:
        # OK: any "Power BI Desktop" window
        if len(pbi_all) > 1:
            titles = [t for _, t in pbi_all]
            print(f"  WARNING: {len(pbi_all)} Power BI Desktop windows found and "
                  f"could not match by report name — using first one.")
            print(f"  Titles: {titles}")
        pbi_hwnd = pbi_all[0][0]
    else:
        print("  INFO: Power BI Desktop window not found — screenshots unavailable")
        print("        Open the .pbip in Power BI Desktop, then re-run --prepare")
        return {}

    matched_title = win32gui.GetWindowText(pbi_hwnd)
    # Use ascii-safe title for printing (Windows terminal may not support all chars)
    safe_title = matched_title.encode('ascii', errors='replace').decode('ascii')
    print(f"  Using window: \"{safe_title}\" (HWND={pbi_hwnd})")

    import time as _time
    import win32process

    # --- 2. Save original window placement so we can restore it afterwards ---
    original_placement = win32gui.GetWindowPlacement(pbi_hwnd)

    # --- 3. Maximise, make topmost, bring to foreground ---
    # HWND_TOPMOST prevents other windows (Teams, etc.) from overlapping PBI
    # Desktop during capture, ensuring screenshots are clean.
    win32gui.ShowWindow(pbi_hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetWindowPos(pbi_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    # Attempt to bring PBI Desktop to the top visually.
    # SetForegroundWindow may fail on Windows 11 (returns 0 silently) but
    # BringWindowToTop is sufficient since UIAutomation select() does not
    # require the window to have keyboard focus.
    try:
        win32gui.SetForegroundWindow(pbi_hwnd)
    except Exception:
        pass
    try:
        win32gui.BringWindowToTop(pbi_hwnd)
    except Exception:
        pass
    _time.sleep(1.0)

    # --- 4. Get window bounds ---
    rect = win32gui.GetWindowRect(pbi_hwnd)
    wl, wt, wr, wb = rect
    print(f"  Window: {wr - wl}x{wb - wt} at ({wl},{wt})")

    # --- 4b. Detect per-window DPI for accurate constant scaling ---
    try:
        _dpi = ctypes.windll.user32.GetDpiForWindow(pbi_hwnd)
        _dpi_factor = _dpi / 96.0
    except Exception:
        _dpi_factor = 1.0
    print(f"  DPI factor: {_dpi_factor:.2f} ({int(_dpi_factor * 100)}%)")

    # --- 5. Determine canvas bounds ---
    # Constants are scaled by _dpi_factor because PrintWindow renders the window
    # at physical pixel resolution (DPI-aware app), so crop offsets must be in
    # physical pixels too.  At 100% DPI factor=1.0; at 150% factor=1.5.
    RIBBON_H    = int(175 * _dpi_factor)  # title-bar + menu + expanded ribbon (175 = empirical base at 96-DPI; at 150% DPI → 262, just past the separator at window-y≈258)
    TABS_H      = int(52  * _dpi_factor)  # page-tab strip + status bar
    LEFT_TRIM   = max(0, -wl)             # off-screen border (maximised window)
    LEFT_RAIL_W = int(48  * _dpi_factor)  # PBI Desktop left navigation rail

    CANVAS_LEFT   = wl + LEFT_TRIM + LEFT_RAIL_W
    CANVAS_TOP    = wt + RIBBON_H
    CANVAS_BOTTOM = wb - TABS_H
    CANVAS_RIGHT  = wr       # tightened below if right panel detected

    # Clamp canvas bounds to the visible screen area.
    # A maximised window often has its frame slightly off-screen (e.g. x=-11),
    # so wl/wr can be negative or > screen width — ImageGrab fails in that case.
    _screen_w = ctypes.windll.user32.GetSystemMetrics(0)
    _screen_h = ctypes.windll.user32.GetSystemMetrics(1)
    CANVAS_LEFT   = max(0, CANVAS_LEFT)
    CANVAS_TOP    = max(0, CANVAS_TOP)
    CANVAS_RIGHT  = min(CANVAS_RIGHT,  _screen_w)
    CANVAS_BOTTOM = min(CANVAS_BOTTOM, _screen_h)

    # --- 5b. Helpers for window capture (defined early; used by panel detection too) ---

    def _pw_full_image():
        """Capture the full PBI Desktop window via PrintWindow(PW_RENDERFULLCONTENT=2).

        PrintWindow bypasses the screen compositor and works reliably on Windows 11
        regardless of foreground/focus state.
        Returns (PIL.Image, win_rect) where win_rect is (wl, wt, wr, wb).
        """
        import win32ui as _win32ui
        _win_rect = win32gui.GetWindowRect(pbi_hwnd)
        _ww  = _win_rect[2] - _win_rect[0]
        _wh  = _win_rect[3] - _win_rect[1]
        _hdc = win32gui.GetWindowDC(pbi_hwnd)
        _dc  = _win32ui.CreateDCFromHandle(_hdc)
        _mem = _dc.CreateCompatibleDC()
        _bmp = _win32ui.CreateBitmap()
        _bmp.CreateCompatibleBitmap(_dc, _ww, _wh)
        _mem.SelectObject(_bmp)
        ctypes.windll.user32.PrintWindow(pbi_hwnd, _mem.GetSafeHdc(), 2)
        _info = _bmp.GetInfo()
        _data = _bmp.GetBitmapBits(True)
        _full = Image.frombuffer(
            'RGB', (_info['bmWidth'], _info['bmHeight']),
            _data, 'raw', 'BGRX', 0, 1,
        )
        # Cleanup: memory DC first, then release window DC (not DeleteDC), then bitmap
        _mem.DeleteDC()
        win32gui.ReleaseDC(pbi_hwnd, _hdc)
        win32gui.DeleteObject(_bmp.GetHandle())
        return _full, _win_rect

    def _grab_canvas():
        _full, _win_rect = _pw_full_image()
        _cl = CANVAS_LEFT   - _win_rect[0]
        _ct = CANVAS_TOP    - _win_rect[1]
        _cr = CANVAS_RIGHT  - _win_rect[0]
        _cb = CANVAS_BOTTOM - _win_rect[1]
        return _full.crop((_cl, _ct, _cr, _cb))

    import hashlib as _hashlib
    def _hash_img(img):
        return _hashlib.md5(img.tobytes()).hexdigest()[:8]

    def _trim_workspace_gray(img, ws_color=224, tol=2.0, max_trim=80):
        """Remove residual PBI workspace gray from all 4 edges of a captured
        page image.  A row/col is trimmed if its mean is within *tol* of
        *ws_color* AND its std < 1.0 (perfectly uniform gray).  Stops as
        soon as a non-gray row/col is found (no interior scanning).

        *max_trim* caps how many pixels can be removed per edge to avoid
        trimming gray report content (e.g. header bands).
        """
        try:
            import numpy as _np2
        except ImportError:
            return img
        gray = _np2.array(img.convert('L'), dtype=_np2.float32)
        h, w = gray.shape
        t = b = l = r = 0
        for row in range(min(max_trim, h // 4)):
            if gray[row, :].std() < 1.0 and abs(float(gray[row, :].mean()) - ws_color) <= tol:
                t += 1
            else:
                break
        for row in range(h - 1, max(h - max_trim - 1, h * 3 // 4), -1):
            if gray[row, :].std() < 1.0 and abs(float(gray[row, :].mean()) - ws_color) <= tol:
                b += 1
            else:
                break
        for col in range(min(max_trim, w // 4)):
            if gray[:, col].std() < 1.0 and abs(float(gray[:, col].mean()) - ws_color) <= tol:
                l += 1
            else:
                break
        for col in range(w - 1, max(w - max_trim - 1, w * 3 // 4), -1):
            if gray[:, col].std() < 1.0 and abs(float(gray[:, col].mean()) - ws_color) <= tol:
                r += 1
            else:
                break
        if t or b or l or r:
            return img.crop((l, t, w - r, h - b))
        return img

    def _crop_by_page_aspect(img, page_width, page_height):
        """Compute the exact report rectangle using the known page aspect
        ratio from the PBIP definition.  PBI Desktop's 'Fit to page' mode
        centres the report on a gray workspace — knowing the aspect ratio
        lets us find the report rect deterministically without scanning
        pixel variance.

        Returns (left, top, right, bottom) in PIL crop convention
        (right/bottom exclusive), or None if dimensions are unavailable.
        """
        if not page_width or not page_height:
            return None
        canvas_w, canvas_h = img.size
        aspect = page_width / page_height
        # Determine whether the gray bars are horizontal or vertical
        if canvas_w / canvas_h > aspect:
            # Canvas wider than page → vertical gray bars on left/right
            report_h = canvas_h
            report_w = round(report_h * aspect)
            offset_x = (canvas_w - report_w) // 2
            offset_y = 0
        else:
            # Canvas taller than page → horizontal gray bars on top/bottom
            report_w = canvas_w
            report_h = round(report_w / aspect)
            offset_x = 0
            offset_y = (canvas_h - report_h) // 2
        # Sanity: result must be reasonably sized
        if report_w < 200 or report_h < 150:
            return None
        return (offset_x, offset_y, offset_x + report_w, offset_y + report_h)

    def _detect_report_rect(img):
        """Detect the PBI report page rectangle within the rough canvas crop.

        PBI Desktop in 'Fit to page' mode places the report on a uniform gray
        workspace and draws a dotted border around it.  Scan inward from each
        edge: workspace rows/cols have very low std-dev (~0–2), the first
        row/col with noticeably higher std is the dotted border / content.

        The workspace baseline std is estimated from the lowest per-edge
        minimum — at least one edge always has clean workspace pixels.

        Returns (left, top, right, bottom) using PIL crop convention (right
        and bottom are EXCLUSIVE), or None if detection fails.
        Does NOT assume any report background colour.
        """
        try:
            import numpy as _np
        except ImportError:
            return None

        gray = _np.array(img.convert('L'), dtype=_np.float32)
        h, w = gray.shape
        if h < 100 or w < 100:
            return None

        row_std = gray.std(axis=1)
        col_std = gray.std(axis=0)

        # Workspace baseline: take the minimum std from each edge's outer
        # 10 rows/cols, then use the overall minimum.  At least one edge
        # (typically bottom or left) will be clean workspace.
        _edge = 10
        _edge_mins = []
        if h > 2 * _edge:
            _edge_mins.append(float(row_std[:_edge].min()))
            _edge_mins.append(float(row_std[-_edge:].min()))
        if w > 2 * _edge:
            _edge_mins.append(float(col_std[:_edge].min()))
            _edge_mins.append(float(col_std[-_edge:].min()))
        if not _edge_mins:
            return None

        _ws_std = min(_edge_mins)
        _VT = max(_ws_std + 3.0, 3.5)

        def _scan_inward(arr, from_end=False):
            """Return the index of the first element > _VT, scanning from
            the start (from_end=False) or from the end (from_end=True)."""
            indices = range(len(arr) - 1, -1, -1) if from_end else range(len(arr))
            for idx in indices:
                if arr[idx] > _VT:
                    return idx
            return None

        # Bottom edge (scan upward) — most important for the cropping bug
        bottom = _scan_inward(row_std, from_end=True)
        top    = _scan_inward(row_std, from_end=False)
        left   = _scan_inward(col_std, from_end=False)
        right  = _scan_inward(col_std, from_end=True)

        if any(v is None for v in (top, bottom, left, right)):
            return None

        # Skip past the dotted-border dashes AND any remaining uniform-gray
        # workspace gap between the border and the actual report content.
        # The border itself is 1–3 px, but PBI Desktop may add another 3–6 px
        # of uniform gray padding.  Instead of a hardcoded skip, we scan
        # inward from the border until we hit a row/col that is NOT uniform
        # gray (either has content variation or differs from workspace color).
        row_mean = gray.mean(axis=1)
        col_mean = gray.mean(axis=0)

        # Workspace gray is the mean of the outermost clean rows/cols
        _ws_mean = float(_np.median(_np.concatenate([
            row_mean[:min(5, top)], col_mean[:min(5, left)]
        ]))) if top > 0 or left > 0 else 224.0

        def _skip_uniform_gray(arr_std, arr_mean, start, direction=1, limit=60):
            """Advance from *start* in *direction* (+1 or -1) while rows/cols
            are uniform gray (low std AND mean close to workspace gray).
            Returns the first index that is NOT uniform gray.
            limit=60 handles up to ~200 % DPI scaling gaps safely; the
            std and mean guards prevent over-cropping into report content."""
            pos = start
            for _ in range(limit):
                next_pos = pos + direction
                if next_pos < 0 or next_pos >= len(arr_std):
                    break
                if arr_std[next_pos] > _VT:
                    break  # high variance = content
                if abs(float(arr_mean[next_pos]) - _ws_mean) > 8:
                    break  # different color = report background (even if uniform)
                pos = next_pos
            return pos

        top    = _skip_uniform_gray(row_std, row_mean, top,    direction=+1)
        bottom = _skip_uniform_gray(row_std, row_mean, bottom, direction=-1)
        left   = _skip_uniform_gray(col_std, col_mean, left,   direction=+1)
        right  = _skip_uniform_gray(col_std, col_mean, right,  direction=-1)

        # Sanity: result must be reasonably sized
        if (right - left) < 200 or (bottom - top) < 150:
            return None

        return (left, top, right + 1, bottom + 1)

    # --- Detect right panel (Visualizations / Fields / Format) by white-pixel fraction ---
    # Key insight: PBI editing panels have a gray background (~RGB 242) that
    # makes <5% of column pixels reach brightness>=248 (pure white).  The
    # report canvas has lots of white space between chart elements, so >=35%
    # of pixels are bright-white (>=248).  This reliably distinguishes them
    # even when the Visualizations pane contains colorful chart-type icons
    # (which fooled variance-based approaches).
    #
    # Algorithm: scan right-to-left.  A column is "canvas-like" when its
    # white fraction (pixels >=248 / total sampled pixels) >= 0.35.
    # Find the first run of >=20 consecutive canvas-like columns — the
    # rightmost of those columns is the canvas right boundary.
    # Everything to the right = editing panel, cropped out.
    _time.sleep(0.5)
    _cal_full, _cal_wr = _pw_full_image()
    _cal_wl, _cal_wt   = _cal_wr[0], _cal_wr[1]
    _cal_left_rel   = max(0, CANVAS_LEFT          - _cal_wl)
    _cal_right_rel  = min(_cal_full.width  - 1, CANVAS_RIGHT  - _cal_wl)
    _cal_top_rel    = max(0, CANVAS_TOP + 5        - _cal_wt)
    _cal_bottom_rel = min(_cal_full.height - 1, CANVAS_BOTTOM - _cal_wt)
    if _cal_right_rel > _cal_left_rel and _cal_bottom_rel > _cal_top_rel:
        cal_img  = _cal_full.crop((_cal_left_rel, _cal_top_rel,
                                   _cal_right_rel, _cal_bottom_rel))
        cal_w, cal_h = cal_img.size

        def _col_white_frac(img, x_col, h, threshold=248):
            """Fraction of sampled pixels with brightness >= threshold."""
            bright, total = 0, 0
            for y in range(0, h, 3):
                px = img.getpixel((x_col, y))
                if (px[0] + px[1] + px[2]) // 3 >= threshold:
                    bright += 1
                total += 1
            return bright / total if total > 0 else 0.0

        # Canvas columns: white fraction >= 0.35 (lots of white background)
        # Panel columns:  white fraction <  0.25 (gray bg, ~RGB 242, not bright-white)
        CANVAS_FRAC_THRESHOLD = 0.35   # column is "canvas-like" above this
        CANVAS_STREAK_NEEDED  = 50     # consecutive canvas-like columns to confirm
        # 50 > the ~32px separator between Visualizations and Data panes,
        # so the inter-pane white gap doesn't prematurely stop the scan.
        PANEL_SEARCH_WIDTH    = min(800, cal_w)

        consecutive_canvas = 0
        canvas_right_edge  = None     # rightmost canvas-like column (cal_img coords)

        for dx in range(PANEL_SEARCH_WIDTH):
            x_img = cal_w - 1 - dx
            frac  = _col_white_frac(cal_img, x_img, cal_h)
            if frac >= CANVAS_FRAC_THRESHOLD:
                consecutive_canvas += 1
                if consecutive_canvas >= CANVAS_STREAK_NEEDED:
                    # rightmost of these 20 canvas cols = x_img + 19
                    canvas_right_edge = x_img + (CANVAS_STREAK_NEEDED - 1)
                    break
            else:
                consecutive_canvas = 0    # reset on any panel-like column

        panel_start_img_x = None
        if canvas_right_edge is not None:
            _potential_start  = canvas_right_edge + 1
            _potential_panel_w = cal_w - _potential_start
            _min_panel = int(150 * _dpi_factor)   # minimum plausible editing-panel width
            if _potential_panel_w >= _min_panel:
                panel_start_img_x = _potential_start
                CANVAS_RIGHT = CANVAS_LEFT + panel_start_img_x
                print(f"  Right panel detected ({_potential_panel_w}px) at screen x={CANVAS_RIGHT} -- cropped out")
            else:
                print(f"  Right panel region ({_potential_panel_w}px < {_min_panel}px minimum) -- keeping full width")
        else:
            print(f"  No panel boundary found (full canvas visible or no white-space in scan) -- keeping full width")

    print(f"  Canvas: ({CANVAS_LEFT},{CANVAS_TOP}) -> ({CANVAS_RIGHT},{CANVAS_BOTTOM})"
          f"  [{CANVAS_RIGHT - CANVAS_LEFT}x{CANVAS_BOTTOM - CANVAS_TOP}]")

    # --- 6. Navigate through pages using UIAutomation select() ---
    #
    # pywinauto's UIAutomation backend can call select() on PBI Desktop's
    # TabItem controls without requiring the window to be in the foreground or
    # have keyboard focus.  This is the only reliable navigation method for
    # PBI Desktop on Windows 11 (keyboard shortcuts require foreground focus
    # which cannot be obtained programmatically).
    #
    # select() uses the SelectionItem UIA pattern.  Note: invoke() does NOT work
    # on PBI Desktop's custom WPF tab controls (Invoke pattern is not exposed).
    n = len(pages)
    Path('temp').mkdir(exist_ok=True)
    image_map: dict = {}

    try:
        from pywinauto import Application as _PwaApp
        _pwa_app = _PwaApp(backend='uia').connect(handle=pbi_hwnd)
        _pbi_win = _pwa_app.window(handle=pbi_hwnd)

        # Find the widest Tab control in the page-tab-strip area.
        # The strip sits near the bottom of the window: top ~ screen_h - 75,
        # bottom ~ screen_h - 30, width > 500 px.
        # The tab strip sits just above the window bottom (wb).
        # Use wb-relative bounds so the search works regardless of screen resolution.
        _tab_strip  = None
        _best_width = 0
        _best_total = 0   # total TabItem children — tie-break after width filter
        _tab_candidates = []

        # Step A: ensure PBI Desktop is showing the Report view canvas.
        # The left view-switcher rail is a VERTICAL Tab with items stacked top-to-bottom
        # (Report / Data / Model / DAX / TMDL).  Its height >> its width (~41px wide,
        # ~800px tall).  We detect it by aspect ratio and select item[0] = Report view.
        # NOTE: do NOT match by width < 60 alone — some ribbon groups are also narrow.
        try:
            for _ctrl in _pbi_win.descendants(control_type='Tab'):
                _r  = _ctrl.rectangle()
                _rw = _r.right  - _r.left
                _rh = _r.bottom - _r.top
                # Left nav rail: taller than it is wide, anchored at left edge, < 100px wide
                if _rw < 100 and _rh > _rw * 3 and _r.left < (wl + 100):
                    _vtabs = _ctrl.children(control_type='TabItem')
                    if _vtabs:
                        _vtabs[0].select()   # Report view = first item
                        _time.sleep(0.8)
                        print("  Ensured Report view is active")
                        break
        except Exception:
            pass   # non-critical; navigation proceeds regardless

        # Step A2: set Page View = "Fit to page" so report content scales to the
        # visible canvas area and nothing overflows the bottom edge.
        # Approach: click View ribbon tab → click "Page view" button to open
        # dropdown → check if "Fit to page" is already toggled on → if not,
        # click it → close dropdown → return to Home tab.
        try:
            _view_tab_ctrl = None
            for _ctrl in _pbi_win.descendants(control_type='TabItem'):
                if _ctrl.window_text().strip().lower() == 'view':
                    _view_tab_ctrl = _ctrl
                    break
            if _view_tab_ctrl:
                _view_tab_ctrl.select()
                _time.sleep(0.8)

                # Click the "Page view" button in the ribbon (y < 300) to
                # open the dropdown menu with Fit to page / Fit to width /
                # Actual size checkboxes.
                _pv_btn = None
                for _ctrl in _pbi_win.descendants(control_type='Button'):
                    _bt = _ctrl.window_text().strip().lower()
                    if _bt == 'page view' and _ctrl.rectangle().top < 300:
                        _pv_btn = _ctrl
                        break

                _ftp_done = False
                if _pv_btn:
                    _pv_btn.click_input()
                    _time.sleep(0.8)
                    # Look for the "Fit to page" CheckBox in the dropdown
                    for _ctrl in _pbi_win.descendants(control_type='CheckBox'):
                        if _ctrl.window_text().strip().lower() == 'fit to page':
                            try:
                                _state = _ctrl.get_toggle_state()
                            except Exception:
                                _state = 0
                            if _state == 1:
                                # Already active — close dropdown with Escape
                                from pywinauto.keyboard import send_keys as _send_keys
                                _send_keys('{ESC}')
                                _time.sleep(0.3)
                                print("  Page View = Fit to page (already active)")
                            else:
                                _ctrl.click_input()
                                _time.sleep(0.5)
                                print("  Set Page View = Fit to page")
                            _ftp_done = True
                            break
                    if not _ftp_done:
                        # Dropdown opened but checkbox not found — close it
                        from pywinauto.keyboard import send_keys as _send_keys
                        _send_keys('{ESC}')
                        _time.sleep(0.3)

                if not _ftp_done:
                    # Fallback: look for a Button/SplitButton with "fit to page"
                    # anywhere in the ribbon area (y < 300).
                    for _ctrl in _pbi_win.descendants(control_type='Button'):
                        _bt = _ctrl.window_text().strip().lower()
                        if 'fit to page' in _bt and _ctrl.rectangle().top < 300:
                            _ctrl.click_input()
                            _time.sleep(0.3)
                            _ftp_done = True
                            print("  Set Page View = Fit to page (fallback)")
                            break

                # Return to Home tab regardless of success
                for _ctrl in _pbi_win.descendants(control_type='TabItem'):
                    if _ctrl.window_text().strip().lower() == 'home':
                        _ctrl.select()
                        _time.sleep(0.3)
                        break
        except Exception as _ftp_err:
            print(f"  Page view setup skipped: {_ftp_err}")

        # Step B: find the bottom page tab strip.
        # Two hard constraints that uniquely identify it:
        #   1. TOP is in the bottom 20% of the window (near wb) — excludes ribbon tabs
        #   2. WIDTH > 500px — excludes left/right navigation rails
        # Among controls that pass both, prefer the one with the most TabItem children.
        _bottom_zone = wt + int((wb - wt) * 0.80)   # top 80% of window = ribbon area
        for _ctrl in _pbi_win.descendants(control_type='Tab'):
            _r  = _ctrl.rectangle()
            _rw = _r.right - _r.left
            _tab_candidates.append((_r.left, _r.top, _r.right, _r.bottom, _rw))

            if _rw < 500:           # too narrow — not the full-width page strip
                continue
            if _r.top < _bottom_zone:  # in the top 80% — ribbon/toolbar, not page tabs
                continue

            try:
                _total = len(_ctrl.children(control_type='TabItem'))
            except Exception:
                _total = 0

            if _total > _best_total or (_total == _best_total and _rw > _best_width):
                _tab_strip  = _ctrl
                _best_width = _rw
                _best_total = _total

        # Fallback A: relax the width to 200px (unusual DPI or small window)
        if _tab_strip is None:
            for _ctrl in _pbi_win.descendants(control_type='Tab'):
                _r  = _ctrl.rectangle()
                _rw = _r.right - _r.left
                if _rw < 200 or _r.top < _bottom_zone:
                    continue
                try:
                    _total = len(_ctrl.children(control_type='TabItem'))
                except Exception:
                    _total = 0
                if _total > _best_total or (_total == _best_total and _rw > _best_width):
                    _tab_strip  = _ctrl
                    _best_width = _rw
                    _best_total = _total

        # Fallback B: original position-only heuristic (wb-relative, any width)
        if _tab_strip is None:
            _best_width = 0
            for _ctrl in _pbi_win.descendants(control_type='Tab'):
                _r  = _ctrl.rectangle()
                _rw = _r.right - _r.left
                if (_r.top > (wb - 100) and _r.bottom < (wb + 10) and _rw > _best_width):
                    _tab_strip  = _ctrl
                    _best_width = _rw

        if _tab_strip is None:
            print(f"  DEBUG: {len(_tab_candidates)} Tab control(s) found (wb={wb}):")
            for _tc in _tab_candidates[:10]:
                print(f"    rect=({_tc[0]},{_tc[1]},{_tc[2]},{_tc[3]}) w={_tc[4]}")
            raise RuntimeError("Page tab strip not found via UIAutomation")

        _r0 = _tab_strip.rectangle()
        print(f"  UIAutomation tab strip: rect=({_r0.left},{_r0.top},"
              f"{_r0.right},{_r0.bottom})  w={_r0.right - _r0.left}")

        # Tighten CANVAS_BOTTOM to the actual tab strip top — more accurate than
        # the wb - TABS_H estimate which can cut off chart content at the bottom.
        if _r0.top > CANVAS_TOP + 100:
            CANVAS_BOTTOM = min(CANVAS_BOTTOM, _r0.top)
            print(f"  CANVAS_BOTTOM tightened to tab strip top: {CANVAS_BOTTOM}")

        _page_tabs = _tab_strip.children(control_type='TabItem')
        n_tabs = len(_page_tabs)
        print(f"  Found {n_tabs} TabItem controls")

        if n_tabs == 0:
            raise RuntimeError("No TabItem children found in tab strip")

        # Match pages list -> tab indices by display_name (normalised).
        # Fall back to positional order for unmatched pages.
        def _norm_name(s: str) -> str:
            return re.sub(r'[^\w ]', '', s.lower()).strip()

        _tab_names = []
        for _t in _page_tabs:
            try:
                _tab_names.append(_norm_name(_t.window_text()))
            except Exception:
                _tab_names.append('')

        _tab_for_page: dict = {}
        _used_tabs: set     = set()
        for _pi, _page in enumerate(pages):
            _pn = _norm_name(_page['display_name'])
            for _ti, _tn in enumerate(_tab_names):
                if _pn and _tn and _pn == _tn and _ti not in _used_tabs:
                    _tab_for_page[_pi] = _ti
                    _used_tabs.add(_ti)
                    break

        # Positional fallback for any unmatched pages
        _free_tabs = [_ti for _ti in range(n_tabs) if _ti not in _used_tabs]
        for _pi in range(len(pages)):
            if _pi not in _tab_for_page and _free_tabs:
                _tab_for_page[_pi] = _free_tabs.pop(0)

        # --- 6b. Refine canvas bounds via dotted-border detection ---
        # PBI Desktop's 'Fit to page' mode renders the report on a gray
        # workspace.  Detect the report boundary by scanning pixel variance
        # inward from the canvas edges.
        #
        # NOTE: PBI Desktop does NOT preserve the declared page aspect ratio
        # in the viewport — it stretches the page to fill the available width.
        # Therefore we CANNOT rely on page.json width/height to calculate the
        # report rect.  Page dimensions are stored in the page dicts for
        # downstream use (pbip_context.json, future per-visual cropping) but
        # the screen crop must use pixel-variance detection.
        #
        # IMPORTANT: We only refine LEFT, TOP, and RIGHT from the detected
        # rectangle.  CANVAS_BOTTOM is kept at the tab-strip top (from
        # UIAutomation) because:
        #   1. Different report pages may have content extending to different
        #      depths — the detection runs on just ONE page.
        #   2. A few rows of gray workspace at the bottom is harmless;
        #      clipping actual chart/table content is not.
        _time.sleep(1.0)  # ensure page fully rendered after Fit-to-page
        _rough_img = _grab_canvas()
        _report_rect = _detect_report_rect(_rough_img)

        if _report_rect is not None:
            _rl, _rt, _rr, _rb = _report_rect
            _old_cl, _old_ct = CANVAS_LEFT, CANVAS_TOP
            CANVAS_LEFT   = _old_cl + _rl
            CANVAS_TOP    = _old_ct + _rt
            CANVAS_RIGHT  = _old_cl + _rr
            # Keep CANVAS_BOTTOM unchanged — don't tighten bottom edge
            print(f"  Dotted-border detected: report rect ({_rr - _rl}x{_rb - _rt})"
                  f" at offset ({_rl},{_rt}) in rough canvas")
            print(f"  Canvas refined (LTR only): ({CANVAS_LEFT},{CANVAS_TOP})"
                  f" -> ({CANVAS_RIGHT},{CANVAS_BOTTOM})"
                  f"  [{CANVAS_RIGHT - CANVAS_LEFT}x{CANVAS_BOTTOM - CANVAS_TOP}]")
        else:
            print(f"  Dotted-border detection: not found — using heuristic canvas bounds")

        # Capture a baseline hash BEFORE navigating page 1 so we can detect
        # whether page 1's content has actually changed from whatever PBI was
        # showing before the capture loop started.
        _baseline_hash = _hash_img(_grab_canvas())

        prev_hash = _baseline_hash
        for i, page in enumerate(pages):
            slide_num   = i + 1
            output_path = f"temp/page_{slide_num}.png"

            _ti = _tab_for_page.get(i, i)
            if _ti < n_tabs:
                _page_tabs[_ti].select()
            _time.sleep(2.5)   # wait for page content to fully render

            curr_img  = _grab_canvas()
            curr_hash = _hash_img(curr_img)

            # If hash unchanged from previous page, the navigation hasn't taken
            # effect yet — retry with a longer wait.
            if curr_hash == prev_hash and _ti < n_tabs:
                print(f"    [{slide_num}/{n}] nav stalled -- retrying select()...")
                _page_tabs[_ti].select()
                _time.sleep(3.0)
                curr_img  = _grab_canvas()
                curr_hash = _hash_img(curr_img)

            # Gray-screen guard: detect if PBI Desktop drifted into an editor view
            # (TMDL / DAX / Data / Model view).  Editor backgrounds are a very uniform
            # mid-gray (~229,229,229).  Sample the centre quarter of the canvas;
            # if mean ≥ 210 AND std < 20, the canvas is almost certainly not a dashboard.
            try:
                import numpy as _np
                _cw, _ch = curr_img.size
                _region = curr_img.crop((_cw//4, _ch//4, 3*_cw//4, 3*_ch//4))
                _arr    = _np.array(_region.convert('L'), dtype=float)
                if _arr.mean() >= 210 and _arr.std() < 20:
                    print(f"    [{slide_num}/{n}] editor view detected — resetting to Report view")
                    try:
                        for _vc in _pbi_win.descendants(control_type='Tab'):
                            _rvr = _vc.rectangle()
                            _rvw = _rvr.right - _rvr.left
                            _rvh = _rvr.bottom - _rvr.top
                            if _rvw < 100 and _rvh > _rvw * 3 and _rvr.left < (wl + 100):
                                _vt = _vc.children(control_type='TabItem')
                                if _vt:
                                    _vt[0].select()
                                    _time.sleep(2.0)
                                    break
                    except Exception:
                        pass
                    if _ti < n_tabs:
                        _page_tabs[_ti].select()
                        _time.sleep(2.5)
                    curr_img  = _grab_canvas()
                    curr_hash = _hash_img(curr_img)
            except Exception:
                pass   # numpy unavailable or other error — skip guard

            curr_img = _trim_workspace_gray(curr_img)
            curr_img.save(output_path)
            image_map[slide_num] = output_path
            prev_hash = curr_hash
            safe_name = page['display_name'].encode('ascii', errors='replace').decode('ascii')
            print(f"    [{slide_num}/{n}] {safe_name[:55]}")

    except ImportError:
        print("  WARNING: pywinauto not available -- page navigation requires it")
        print("           Install with: pip install pywinauto")
    except Exception as _nav_err:
        print(f"  WARNING: UIAutomation navigation error: {_nav_err}")

    # --- 8. Remove topmost flag and restore original window state ---
    win32gui.SetWindowPos(pbi_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    # Never restore to minimised -- that would move the window to (-32000,-32000)
    _orig_show = original_placement[1]
    if _orig_show in (2, 6, 11):   # SW_SHOWMINIMIZED, SW_MINIMIZE, SW_FORCEMINIMIZE
        _orig_show = win32con.SW_SHOWNORMAL
    win32gui.ShowWindow(pbi_hwnd, _orig_show)

    print(f"  OK Captured {len(image_map)} page screenshots from Power BI Desktop")
    return image_map


# ---------------------------------------------------------------------------
# 4b. Export PDF from Power BI Desktop via UI Automation
# ---------------------------------------------------------------------------

def _export_pdf_from_pbi_desktop(pbip_path: Path, pbip_root: Path,
                                  n_pages: int, pbip_stem: str = '') -> dict:
    """
    Export a PDF from Power BI Desktop via File > Export > Export to PDF,
    then extract page images using fitz.

    Steps:
      1. Find PBI Desktop window and bring to foreground
      2. Click File tab to open backstage
      3. Click Export tab, then "Export to PDF" button
      4. Wait for PDF generation to complete
      5. Handle any Adobe Acrobat font warning dialogs
      6. Copy PDF from PBI Desktop temp folder (no Save As needed)
      7. Close Adobe Acrobat Reader
      8. Extract page images via fitz

    Returns dict: slide_number (1-based) -> image path string,
    or empty dict if export fails.
    """
    try:
        import ctypes
        import os
        import time
        import shutil
        import win32gui
        import win32con
        import win32api
        import win32process
        import pywinauto
        from pywinauto.keyboard import send_keys
        import fitz
        from PIL import Image as PILImage
        import io as _io
        from ctypes import wintypes
    except ImportError as e:
        print(f"  WARNING: PDF export unavailable — missing dependency: {e}")
        return {}

    user32 = ctypes.windll.user32

    # --- SendInput structures for physical key presses ---
    _INPUT_KEYBOARD = 1
    _KEYEVENTF_KEYUP = 0x0002

    class _KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class _INPUT(ctypes.Structure):
        class _U(ctypes.Union):
            _fields_ = [("ki", _KEYBDINPUT)]
        _fields_ = [
            ("type", wintypes.DWORD),
            ("_input", _U),
        ]

    def _send_key(vk, down=True, up=True):
        inputs = []
        extra = ctypes.pointer(ctypes.c_ulong(0))
        if down:
            ki = _KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=extra)
            inp = _INPUT(type=_INPUT_KEYBOARD)
            inp._input.ki = ki
            inputs.append(inp)
        if up:
            ki = _KEYBDINPUT(wVk=vk, wScan=0, dwFlags=_KEYEVENTF_KEYUP, time=0, dwExtraInfo=extra)
            inp = _INPUT(type=_INPUT_KEYBOARD)
            inp._input.ki = ki
            inputs.append(inp)
        arr = (_INPUT * len(inputs))(*inputs)
        user32.SendInput(len(inputs), arr, ctypes.sizeof(_INPUT))

    def _send_escape():
        _send_key(0x1B)  # VK_ESCAPE

    def _find_window_by_title(keywords: tuple[str, ...]) -> int:
        """Return first visible window whose title contains any keyword."""
        found_hwnd = 0

        def _enum(hwnd, _):
            nonlocal found_hwnd
            if found_hwnd:
                return False
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = (win32gui.GetWindowText(hwnd) or '').lower()
            if not title:
                return True
            if any(k in title for k in keywords):
                found_hwnd = hwnd
                return False
            return True

        win32gui.EnumWindows(_enum, None)
        return found_hwnd

    def _dismiss_acrobat_close_tabs_dialog() -> None:
        """Dismiss Acrobat 'Close all tabs' prompts by choosing Yes."""
        def _enum_dialog(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != '#32770':
                return True
            title = (win32gui.GetWindowText(hwnd) or '').lower()
            if 'acrobat' not in title and 'close all tabs' not in title:
                return True
            try:
                d_app = pywinauto.Application(backend='uia').connect(handle=hwnd)
                d_dlg = d_app.window(handle=hwnd)
                yes_btn = d_dlg.child_window(title='Yes', control_type='Button')
                if yes_btn.exists(timeout=1):
                    yes_btn.click_input()
                    return True
                # Fallback: Alt+Y on focused confirmation dialog.
                try:
                    _bring_to_foreground(hwnd)
                except Exception:
                    pass
                send_keys('%y')
            except Exception:
                pass
            return True

        win32gui.EnumWindows(_enum_dialog, None)

    def _close_acrobat_window(hwnd: int) -> None:
        """Close Acrobat window and handle optional close-all-tabs dialog."""
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            return
        time.sleep(0.6)
        _dismiss_acrobat_close_tabs_dialog()
        time.sleep(0.8)

    # --- Helper: bring a window to the foreground reliably ---
    def _bring_to_foreground(hwnd):
        our_tid = win32api.GetCurrentThreadId()
        fg_hwnd = win32gui.GetForegroundWindow()
        fg_tid, _ = win32process.GetWindowThreadProcessId(fg_hwnd)
        tgt_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        user32.AttachThreadInput(our_tid, fg_tid, True)
        user32.AttachThreadInput(our_tid, tgt_tid, True)
        try:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(0.3)
        user32.AttachThreadInput(our_tid, fg_tid, False)
        user32.AttachThreadInput(our_tid, tgt_tid, False)

    # --- 1. Find PBI Desktop window ---
    # Uses the same multi-strategy approach as _capture_pbi_desktop_screenshots:
    #   Priority 1: title contains PBIP stem + "Power BI Desktop"
    #   Priority 2: title contains PBIP stem (newer PBI builds omit suffix)
    #   Priority 3: any "Power BI Desktop" window
    all_windows = []
    pbi_all = []
    stem_matches = []
    _NON_PBI_SUFFIXES = ('- powerpoint', '- word', '- excel', '- outlook',
                         '- notepad', '- visual studio', '- code',
                         '- adobe', 'acrobat')

    def _enum_pbi(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            return True
        if not title:
            return True
        all_windows.append((hwnd, title))
        if 'Power BI Desktop' in title:
            pbi_all.append((hwnd, title))
        return True
    win32gui.EnumWindows(_enum_pbi, None)

    if pbip_stem:
        match_key = re.sub(r'[^\w ]', '', pbip_stem.lower())[:25].strip()
        for hwnd, title in all_windows:
            norm = re.sub(r'[^\w ]', '', title.lower())
            if match_key and match_key in norm:
                title_lower = title.lower()
                if not any(s in title_lower for s in _NON_PBI_SUFFIXES):
                    stem_matches.append((hwnd, title))

    # Pick best candidate
    pbi_hwnd = 0
    stem_and_pbi = [hw for hw, t in stem_matches if 'Power BI Desktop' in t]
    if stem_and_pbi:
        pbi_hwnd = stem_and_pbi[0]
    elif stem_matches:
        pbi_hwnd = stem_matches[0][0]
    elif pbi_all:
        pbi_hwnd = pbi_all[0][0]

    if not pbi_hwnd:
        print("  Power BI Desktop not found for PDF export")
        return {}

    print(f"  PBI Desktop: HWND={pbi_hwnd}")
    try:
        win32gui.ShowWindow(pbi_hwnd, win32con.SW_MAXIMIZE)
        time.sleep(0.2)
    except Exception:
        pass

    # Close stale Acrobat windows from previous runs to avoid false positives
    # when waiting for export completion.
    existing_acrobat_hwnds = set()

    def _collect_acrobat(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == 'AcrobatSDIWindow':
                existing_acrobat_hwnds.add(hwnd)
        return True

    win32gui.EnumWindows(_collect_acrobat, None)
    for h in list(existing_acrobat_hwnds):
        _close_acrobat_window(h)
    if existing_acrobat_hwnds:
        time.sleep(1.5)

    # --- 2. Trigger export ---
    # Primary: click on-screen controls (File -> Export -> Export to PDF).
    # Fallback: keyboard accelerators.
    _bring_to_foreground(pbi_hwnd)
    time.sleep(0.8)
    export_started_at = time.time()

    try:
        app = pywinauto.Application(backend='uia').connect(handle=pbi_hwnd)
        pbi_dlg = app.window(handle=pbi_hwnd)
        win_rect = win32gui.GetWindowRect(pbi_hwnd)

        def _is_in_window(r) -> bool:
            return (
                win_rect[0] <= r.left <= win_rect[2] and
                win_rect[1] <= r.top <= win_rect[3] and
                r.width() > 0 and r.height() > 0
            )

        def _wait_export_started(seconds: int = 60) -> bool:
            for _ in range(seconds):
                acrobat_probe = 0

                def _find_acrobat_probe(hwnd, _):
                    nonlocal acrobat_probe
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        cls = win32gui.GetClassName(hwnd)
                        if cls == 'AcrobatSDIWindow' and '.pdf' in title.lower():
                            acrobat_probe = hwnd
                    return True

                win32gui.EnumWindows(_find_acrobat_probe, None)
                if acrobat_probe and acrobat_probe not in existing_acrobat_hwnds:
                    return True
                if _find_window_by_title(('save as', 'save print output as')):
                    return True
                time.sleep(1)
            return False

        export_triggered = False

        # Path A: on-screen UIA click sequence.
        send_keys('{ESC}')
        time.sleep(0.2)

        file_tab = None
        for d in pbi_dlg.descendants(control_type='TabItem'):
            try:
                ei = d.element_info
                r = d.rectangle()
                name = (ei.name or '').strip().lower()
                auto_id = (getattr(ei, 'automation_id', '') or '').strip().lower()
                if not _is_in_window(r):
                    continue
                if (
                    (auto_id == 'ribbon-file' or name == 'file') and
                    r.left <= (win_rect[0] + 250) and
                    r.top <= (win_rect[1] + 220)
                ):
                    file_tab = d
                    break
            except Exception:
                pass

        if file_tab:
            print("  Debug: File tab located")
            try:
                file_tab.select()
            except Exception:
                file_tab.click_input()
            time.sleep(1.2)
        else:
            print("  Debug: File tab not found")

        export_tab = None
        for d in pbi_dlg.descendants():
            try:
                ei = d.element_info
                r = d.rectangle()
                if not _is_in_window(r):
                    continue
                ct = (ei.control_type or '').strip()
                if (
                    (ei.name or '').strip() == 'Export' and
                    ct in ('TabItem', 'Text', 'Hyperlink', 'Button') and
                    r.left <= (win_rect[0] + 450) and
                    r.top >= (win_rect[1] + 350)
                ):
                    export_tab = d
                    break
            except Exception:
                pass

        if not export_tab:
            # Recovery: force backstage via Alt+F, then search again.
            send_keys('%f')
            time.sleep(1.0)
            for d in pbi_dlg.descendants(control_type='TabItem'):
                try:
                    ei = d.element_info
                    r = d.rectangle()
                    if not _is_in_window(r):
                        continue
                    ct = (ei.control_type or '').strip()
                    if (
                        (ei.name or '').strip() == 'Export' and
                        ct in ('TabItem', 'Text', 'Hyperlink', 'Button') and
                        r.left <= (win_rect[0] + 450) and
                        r.top >= (win_rect[1] + 350)
                    ):
                        export_tab = d
                        break
                except Exception:
                    pass

        if export_tab:
            print("  Debug: Export tab located")
            try:
                export_tab.select()
            except Exception:
                export_tab.click_input()
            time.sleep(1.0)
        else:
            print("  Debug: Export tab not found")

        export_pdf_btn = None
        for _ in range(30):
            for d in pbi_dlg.descendants(control_type='Button'):
                try:
                    ei = d.element_info
                    r = d.rectangle()
                    if not _is_in_window(r):
                        continue
                    ct = (ei.control_type or '').strip()
                    name = (ei.name or '').strip().lower()
                    if (
                        ('export to pdf' in name) and
                        ct == 'Button' and
                        r.left >= (win_rect[0] + 120)
                    ):
                        export_pdf_btn = d
                        break
                except Exception:
                    pass
            if export_pdf_btn:
                break
            time.sleep(0.3)

        if not export_pdf_btn:
            # Recovery: use Export tab accelerator and rescan.
            send_keys('e')
            time.sleep(0.8)
            for _ in range(20):
                for d in pbi_dlg.descendants(control_type='Button'):
                    try:
                        ei = d.element_info
                        r = d.rectangle()
                        if not _is_in_window(r):
                            continue
                        ct = (ei.control_type or '').strip()
                        name = (ei.name or '').strip().lower()
                        if (
                            ('export to pdf' in name) and
                            ct in ('Button', 'Hyperlink', 'Text', 'ListItem') and
                            r.left >= (win_rect[0] + 120)
                        ):
                            export_pdf_btn = d
                            break
                    except Exception:
                        pass
                if export_pdf_btn:
                    break
                time.sleep(0.3)

        if export_pdf_btn:
            print("  Debug: Export to PDF control located")
            export_pdf_btn.click_input()
            export_triggered = _wait_export_started()
        else:
            print("  Debug: Export to PDF control not found")

        # Path B: keyboard fallback.
        if not export_triggered:
            for attempt in range(2):
                send_keys('{ESC}')
                time.sleep(0.2)
                send_keys('%f')
                time.sleep(0.8 if attempt == 0 else 1.2)
                send_keys('e')
                time.sleep(0.5)
                send_keys('p')
                time.sleep(0.7)
                if _wait_export_started():
                    export_triggered = True
                    break

        if not export_triggered:
            print("  Failed to trigger Export to PDF (click + keyboard paths)")
            return {}
    except (Exception, KeyboardInterrupt) as e:
        print(f"  Failed to trigger Export to PDF: {e}")
        return {}

    # --- 4. Wait for PDF generation (poll for up to 5 minutes) ---
    print("  Generating PDF (this may take a minute)...")
    max_wait = 300
    poll_interval = 2
    elapsed = 0
    acrobat_hwnd = 0

    while elapsed < max_wait:
        time.sleep(poll_interval)
        elapsed += poll_interval

        # Check for Acrobat Reader window (PDF auto-opens when done)
        def _find_acrobat(hwnd, _):
            nonlocal acrobat_hwnd
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                if cls == 'AcrobatSDIWindow' and '.pdf' in title.lower():
                    acrobat_hwnd = hwnd
            return True
        win32gui.EnumWindows(_find_acrobat, None)

        if acrobat_hwnd:
            print(f"  PDF generated ({elapsed}s)")
            break

        # Check for and dismiss Adobe Acrobat font warning
        def _find_dialog(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                if cls == '#32770' and 'Adobe Acrobat' in title:
                    try:
                        d_app = pywinauto.Application(backend='uia').connect(handle=hwnd)
                        d_dlg = d_app.window(handle=hwnd)
                        ok_btn = d_dlg.child_window(title='OK', control_type='Button')
                        if ok_btn.exists(timeout=1):
                            ok_btn.click_input()
                            print("  Dismissed Acrobat font warning")
                    except Exception:
                        pass
            return True
        win32gui.EnumWindows(_find_dialog, None)

    if not acrobat_hwnd:
        print("  Timed out waiting for PDF export")
        return {}

    time.sleep(1)

    # --- 5. Dismiss any remaining Acrobat font warnings ---
    def _dismiss_warnings(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if cls == '#32770' and 'Acrobat' in title:
                try:
                    d_app = pywinauto.Application(backend='uia').connect(handle=hwnd)
                    d_dlg = d_app.window(handle=hwnd)
                    ok_btn = d_dlg.child_window(title='OK', control_type='Button')
                    if ok_btn.exists(timeout=1):
                        ok_btn.click_input()
                except Exception:
                    pass
        return True
    win32gui.EnumWindows(_dismiss_warnings, None)
    time.sleep(1)

    # --- 6. Copy PDF from PBI Desktop temp folder ---
    # PBI Desktop writes exported PDFs to:
    #   %LOCALAPPDATA%\Temp\Power BI Desktop\print-job-<guid>\<filename>.pdf
    # This is far more reliable than automating Acrobat's Save As dialog.
    acr_title = win32gui.GetWindowText(acrobat_hwnd)
    pdf_name = acr_title.split(' - Adobe')[0].strip() if ' - Adobe' in acr_title else ''
    if not pdf_name.lower().endswith('.pdf'):
        pdf_name = pdf_name + '.pdf'

    pbi_temp_dir = Path(os.environ.get('LOCALAPPDATA', '')) / 'Temp' / 'Power BI Desktop'
    target_pdf = pbip_root / pdf_name

    source_pdf = None
    print(f"  Looking for exported PDF in PBI Desktop temp folder...")

    # Wait for the temp PDF to appear (PBI writes it before Acrobat opens,
    # but there can be a small race).
    for attempt in range(30):
        if pbi_temp_dir.exists():
            candidates = []
            for print_job_dir in pbi_temp_dir.glob('print-job-*'):
                candidate = print_job_dir / pdf_name
                if candidate.exists() and candidate.stat().st_size > 0:
                    mtime = candidate.stat().st_mtime
                    if mtime >= (export_started_at - 5):
                        candidates.append((mtime, candidate))
            if candidates:
                # Pick the newest one
                candidates.sort(reverse=True)
                source_pdf = candidates[0][1]
                break
        time.sleep(0.5)

    if not source_pdf:
        print(f"  WARNING: Could not find exported PDF in PBI Desktop temp folder")
        # --- Close Acrobat before returning ---
        _close_acrobat_window(acrobat_hwnd)
        return {}

    print(f"  Found temp PDF: {source_pdf.name} ({source_pdf.stat().st_size:,} bytes)")

    # --- 7. Close Acrobat Reader (must release file lock before copy) ---
    _close_acrobat_window(acrobat_hwnd)
    time.sleep(1)

    # Copy the PDF to the PBIP directory
    copied = False
    last_err = None
    for _ in range(8):
        try:
            shutil.copy2(str(source_pdf), str(target_pdf))
            copied = True
            break
        except Exception as e:
            last_err = e
            _close_acrobat_window(acrobat_hwnd)
            time.sleep(0.8)

    if copied:
        print(f"  PDF saved: {target_pdf.name} ({target_pdf.stat().st_size:,} bytes)")
    elif target_pdf.exists() and target_pdf.stat().st_size > 0:
        print(f"  PDF already at: {target_pdf.name}")
    else:
        print(f"  WARNING: Could not copy PDF: {last_err}")
        # Fall back to reading directly from temp location
        target_pdf = source_pdf

    # --- 8. Extract page images via fitz ---
    try:
        pdf_doc = fitz.open(str(target_pdf))
        zoom = 150 / 72
        mat = fitz.Matrix(zoom, zoom)
        image_map: dict = {}

        Path('temp').mkdir(exist_ok=True)
        for page_idx in range(min(len(pdf_doc), n_pages)):
            page = pdf_doc[page_idx]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = PILImage.open(_io.BytesIO(pix.tobytes("png")))

            img_path = f"temp/page_{page_idx + 1}.png"
            img.save(img_path)
            image_map[page_idx + 1] = img_path

        pdf_doc.close()
        print(f"  OK Extracted {len(image_map)} page images from exported PDF")
        return image_map

    except Exception as e:
        print(f"  WARNING: Failed to extract images from PDF: {e}")
        return {}


# ---------------------------------------------------------------------------
# 5. Companion image extraction (fallback: user-provided PDF/PPTX)
# ---------------------------------------------------------------------------

def _find_companion_export(pbip_path: Path, pbip_root: Path) -> Optional[Path]:
    """
    Look for a PDF or PPTX companion export in the same folder as the .pbip.

    Search order (most specific first):
      1. PDF whose stem matches the .pbip stem exactly
      2. PPTX whose stem matches the .pbip stem exactly
      3. PDF whose stem starts with the .pbip stem
      4. PPTX whose stem starts with the .pbip stem
      5. Any PDF in the folder (excluding outputs we generated)
      6. Any PPTX in the folder (excluding outputs we generated)
    """
    pbip_stem = pbip_path.stem if pbip_path.suffix.lower() == '.pbip' else pbip_root.name

    # Exclude files we generate ourselves
    def is_our_output(p: Path) -> bool:
        return p.stem.endswith('_executive') or p.stem.endswith('-Executive-Insights')

    # Use the first 15 characters of the PBIP stem as a loose match key
    # (e.g. "AI-in-One Dashb" for "AI-in-One Dashboard 1802 - w Agent 365...")
    match_prefix = pbip_stem.lower()[:15]

    for ext in ('.pdf', '.pptx'):
        # 1. Exact name match
        exact = pbip_root / f"{pbip_stem}{ext}"
        if exact.exists() and not is_our_output(exact):
            return exact

        # 2. Prefix match — file whose stem begins with the same words
        for f in sorted(pbip_root.glob(f"*{ext}")):
            if not is_our_output(f) and f.stem.lower().startswith(match_prefix):
                return f

    # No related file found — do NOT fall back to unrelated files
    return None


def _extract_companion_images(pbip_path: Path, pbip_root: Path,
                               n_pages: int) -> dict:
    """
    Find a companion PDF/PPTX, extract images from it, and return a
    dict mapping slide_number (1-based) -> image_path string.

    Skips the first slide of PPTX exports (title/cover page convention),
    includes all pages for PDFs.
    """
    companion = _find_companion_export(pbip_path, pbip_root)
    if companion is None:
        return {}

    ext = companion.suffix.lower()
    print(f"  Using companion: {companion.name}")

    Path('temp').mkdir(exist_ok=True)
    image_map: dict = {}

    try:
        if ext == '.pdf':
            import fitz
            from PIL import Image as PILImage
            import io as _io
            pdf_doc = fitz.open(str(companion))
            zoom = 150 / 72
            mat = fitz.Matrix(zoom, zoom)
            slide_idx = 0
            for page_idx in range(len(pdf_doc)):
                page = pdf_doc[page_idx]
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = PILImage.open(_io.BytesIO(pix.tobytes("png")))
                img_path = f"temp/pbip_page_{page_idx + 1}.png"
                img.save(img_path)
                slide_idx += 1
                if slide_idx <= n_pages:
                    image_map[slide_idx] = img_path
                if slide_idx >= n_pages:
                    break
            pdf_doc.close()

        elif ext == '.pptx':
            from pptx import Presentation as _Prs
            import io as _io
            prs = _Prs(str(companion))
            slide_idx = 0
            for i, slide in enumerate(prs.slides):
                # Skip first slide (Power BI PPTX exports have a cover page)
                if i == 0:
                    continue
                # Extract embedded image from slide
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # Picture
                        from PIL import Image as PILImage
                        img = PILImage.open(_io.BytesIO(shape.image.blob))
                        img_path = f"temp/pbip_page_{i + 1}.png"
                        img.save(img_path)
                        slide_idx += 1
                        if slide_idx <= n_pages:
                            image_map[slide_idx] = img_path
                        break
                if slide_idx >= n_pages:
                    break

    except Exception as e:
        print(f"  WARNING: Failed to extract companion images: {e}")
        return {}

    return image_map


# ---------------------------------------------------------------------------
# 5. Main entry point
# ---------------------------------------------------------------------------

def prepare_pbip_for_analysis(pbip_path: str) -> str:
    """
    Main entry point. Accepts path to a .pbip file or its parent folder.

    Writes:
      - temp/analysis_request.json   (same schema as PDF/PPTX pipeline)
      - temp/pbip_context.json       (enriched model data + DAX queries)

    Returns path to analysis_request.json.
    """
    print("=" * 70)
    print("PREPARING PBIP PROJECT FOR ANALYSIS")
    print("=" * 70)

    pbip_path = Path(pbip_path)

    # Resolve to parent directory (works whether user passed .pbip file or folder)
    if pbip_path.suffix.lower() == '.pbip':
        pbip_root = pbip_path.parent
    elif pbip_path.is_dir():
        pbip_root = pbip_path
    else:
        raise ValueError(f"Expected a .pbip file or directory, got: {pbip_path}")

    print(f"\nProject root: {pbip_root}")

    def _safe(text: str) -> str:
        """Strip characters the Windows terminal can't encode."""
        return re.sub(r'[^\x00-\x7F\x80-\xFF]', '?', str(text))

    # Step 1: Discover report pages
    # Pass the .pbip file path (not parent dir) so the function can derive
    # the correct .Report folder name when multiple PBIPs share a directory.
    print("\nDiscovering report pages...")
    pages = discover_report_pages(str(pbip_path))
    visible_pages = [p for p in pages if not p['is_hidden']]
    print(f"  Found {len(visible_pages)} visible pages "
          f"({len(pages) - len(visible_pages)} hidden)")
    for p in visible_pages:
        n_visuals = len(p['visuals'])
        print(f"  - {_safe(p['display_name'])} ({n_visuals} visuals)")

    # Step 2: Extract model metadata
    print("\nExtracting semantic model metadata...")
    model = extract_model_metadata(str(pbip_path))
    print(f"  Found {len(model['tables'])} tables, "
          f"{len(model['measures'])} measures, "
          f"{len(model['relationships'])} relationships")

    # Step 3: Build DAX queries
    print("\nBuilding DAX queries for each page...")
    dax_queries = build_dax_queries(visible_pages, model)
    total_queries = sum(len(d['queries']) for d in dax_queries)
    print(f"  Generated {total_queries} queries across {len(dax_queries)} pages")

    # Step 4: Find images for each page
    pbip_stem = pbip_path.stem if pbip_path.suffix.lower() == '.pbip' else pbip_root.name
    page_images = {}

    # Priority 1: Export PDF from PBI Desktop via UI Automation (cleanest output)
    print("\nAttempting PDF export from Power BI Desktop...")
    try:
        page_images = _export_pdf_from_pbi_desktop(
            pbip_path, pbip_root, len(visible_pages), pbip_stem=pbip_stem
        )
    except KeyboardInterrupt:
        print("  PDF export interrupted — falling back to screenshots")
        page_images = {}

    if not page_images:
        # Priority 2: Companion PDF/PPTX already in the folder
        print("  Checking for companion PDF/PPTX export in the same folder...")
        page_images = _extract_companion_images(pbip_path, pbip_root, len(visible_pages))
        if page_images:
            matched = sum(1 for v in page_images.values() if v)
            print(f"  Found images for {matched} of {len(visible_pages)} pages from companion file")

    if not page_images:
        # Priority 3: Live screenshots from PBI Desktop (may have gray border artifacts)
        print("\nFalling back to screenshot capture from Power BI Desktop...")
        print("  (Close the Fields / Visualizations / Filters panels for cleanest results)")
        page_images = _capture_pbi_desktop_screenshots(visible_pages, pbip_stem=pbip_stem)

    if not page_images:
        print("  No images found — slides will be text-only")
        print("  Tip: Open the .pbip in Power BI Desktop, then re-run --prepare to capture live visuals")

    # Step 5: Write output files
    Path('temp').mkdir(exist_ok=True)

    # analysis_request.json — same schema; image_path populated if images found
    slides_meta = []
    for i, page in enumerate(visible_pages, start=1):
        slides_meta.append({
            'slide_number': i,
            'title': page['display_name'],
            'image_path': page_images.get(i),
            'slide_type': _classify_page_type(page['display_name']),
            'source_type': 'pbip',
        })

    request_data = {
        'source_file': str(pbip_path),
        'source_type': 'pbip',
        'total_slides': len(slides_meta),
        'slides': slides_meta,
    }
    request_file = 'temp/analysis_request.json'
    with open(request_file, 'w', encoding='utf-8') as f:
        json.dump(request_data, f, indent=2)

    # pbip_context.json — enriched model data
    context_data = {
        'pbip_path': str(pbip_path),
        'pages': visible_pages,
        'model': model,
        'dax_queries': dax_queries,
    }
    context_file = 'temp/pbip_context.json'
    with open(context_file, 'w', encoding='utf-8') as f:
        json.dump(context_data, f, indent=2)

    print(f"\nOK Analysis request saved to: {request_file}")
    print(f"OK PBIP context saved to:      {context_file}")

    return request_file


def _classify_page_type(display_name: str) -> str:
    """Mirror the classify_slide_type logic from the PPTX/PDF pipelines."""
    name_lower = display_name.lower()
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
