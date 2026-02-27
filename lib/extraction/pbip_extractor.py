"""
PBIP Extraction Module

Extracts report pages, semantic model metadata, and builds DAX queries
for Power BI PBIP projects. Enables Claude to query the live in-memory
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

    # Find the .Report folder
    report_dir = None
    for candidate in root.iterdir() if root.is_dir() else []:
        if candidate.is_dir() and candidate.name.endswith('.Report'):
            report_dir = candidate
            break

    # If pbip_root itself is the .Report folder
    if report_dir is None and root.name.endswith('.Report'):
        report_dir = root

    if report_dir is None:
        # Try one level up
        parent = root.parent
        for candidate in parent.iterdir():
            if candidate.is_dir() and candidate.name.endswith('.Report'):
                report_dir = candidate
                break

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
    model_dir = None
    search_root = root if root.is_dir() else root.parent
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

    # Build a normalised match key from the PBIP stem for fuzzy title matching
    if pbip_stem:
        match_key = re.sub(r'[^\w ]', '', pbip_stem.lower())[:25].strip()
        for hwnd, title in all_windows:
            norm = re.sub(r'[^\w ]', '', title.lower())
            if match_key and match_key in norm:
                stem_matches.append((hwnd, title))

    # Pick the best candidate
    pbi_hwnd = None
    if stem_matches:
        # Best: window title contains the PBIP file name
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

    # --- 5. Determine canvas bounds ---
    RIBBON_H    = 130        # title-bar + menu + quick-access + collapsed ribbon
    TABS_H      = 52         # page-tab strip + status bar
    LEFT_TRIM   = max(0, -wl)  # off-screen border (usually 8px for maximised window)
    LEFT_RAIL_W = 48         # PBI Desktop left navigation rail (collapsed icon bar)

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

    # --- Detect right panel (Visualizations / Fields / Format) by pixel variance ---
    # Canvas content (charts, tables) has HIGH variance; the panel background
    # is nearly uniform (LOW variance per column).
    # Find the leftmost run of >=30 low-variance columns from the right edge.
    _time.sleep(0.5)
    _cal_full, _cal_wr = _pw_full_image()
    _cal_wl, _cal_wt   = _cal_wr[0], _cal_wr[1]
    _cal_left_rel   = max(0, CANVAS_LEFT          - _cal_wl)
    _cal_right_rel  = min(_cal_full.width  - 1, CANVAS_RIGHT  - _cal_wl)
    _cal_top_rel    = max(0, CANVAS_TOP + 100      - _cal_wt)
    _cal_bottom_rel = min(_cal_full.height - 1, CANVAS_TOP + 300 - _cal_wt)
    if _cal_right_rel > _cal_left_rel and _cal_bottom_rel > _cal_top_rel:
        cal_img  = _cal_full.crop((_cal_left_rel, _cal_top_rel,
                                   _cal_right_rel, _cal_bottom_rel))
        cal_w, cal_h = cal_img.size

        def _col_var(img, x_col, h):
            vals = []
            for y in range(0, h, 4):
                px = img.getpixel((x_col, y))
                vals.append((px[0] + px[1] + px[2]) // 3)
            if len(vals) < 2:
                return 9999
            mean = sum(vals) / len(vals)
            return sum((v - mean) ** 2 for v in vals) / len(vals)

        PANEL_SEARCH_WIDTH = min(600, cal_w)
        uniform_streak = 0
        panel_start_img_x = None
        for dx in range(PANEL_SEARCH_WIDTH):
            x_img = cal_w - 1 - dx
            var = _col_var(cal_img, x_img, cal_h)
            if var < 80:
                uniform_streak += 1
            else:
                if uniform_streak >= 30:
                    panel_start_img_x = x_img + 1
                uniform_streak = 0

        if panel_start_img_x is not None:
            CANVAS_RIGHT = CANVAS_LEFT + panel_start_img_x
            print(f"  Right panel detected at screen x={CANVAS_RIGHT} -- cropped out")

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
        _best_named = 0   # count of named TabItem children — primary selection criterion
        _tab_candidates = []
        for _ctrl in _pbi_win.descendants(control_type='Tab'):
            _r  = _ctrl.rectangle()
            _rw = _r.right - _r.left
            _tab_candidates.append((_r.left, _r.top, _r.right, _r.bottom, _rw))

            # Strategy 1 (preferred): Tab with the most named TabItem children.
            # This works even when PBI Desktop reports tab coordinates in WPF
            # logical pixels (DPI-scaled) that don't match screen coordinates.
            try:
                _children = _ctrl.children(control_type='TabItem')
                _named = sum(1 for _c in _children if _c.window_text().strip())
            except Exception:
                _named = 0
            if _named > _best_named or (_named == _best_named and _rw > _best_width):
                _tab_strip  = _ctrl
                _best_width = _rw
                _best_named = _named

        # Strategy 2 fallback: position-based (original logic) if no named children found
        if _best_named == 0:
            _tab_strip = None
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

        prev_hash = None
        for i, page in enumerate(pages):
            slide_num   = i + 1
            output_path = f"temp/page_{slide_num}.png"

            _ti = _tab_for_page.get(i, i)
            if _ti < n_tabs:
                _page_tabs[_ti].select()
            _time.sleep(1.5)   # wait for page content to fully render

            curr_img  = _grab_canvas()
            curr_hash = _hash_img(curr_img)

            # If hash unchanged, retry once with a longer wait
            if prev_hash is not None and curr_hash == prev_hash and _ti < n_tabs:
                print(f"    [{slide_num}/{n}] nav stalled -- retrying select()...")
                _page_tabs[_ti].select()
                _time.sleep(2.0)
                curr_img  = _grab_canvas()
                curr_hash = _hash_img(curr_img)

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

def prepare_pbip_for_claude_analysis(pbip_path: str) -> str:
    """
    Main entry point. Accepts path to a .pbip file or its parent folder.

    Writes:
      - temp/analysis_request.json   (same schema as PDF/PPTX pipeline)
      - temp/pbip_context.json       (enriched model data + DAX queries)

    Returns path to analysis_request.json.
    """
    print("=" * 70)
    print("PREPARING PBIP PROJECT FOR CLAUDE ANALYSIS")
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
    print("\nDiscovering report pages...")
    pages = discover_report_pages(str(pbip_root))
    visible_pages = [p for p in pages if not p['is_hidden']]
    print(f"  Found {len(visible_pages)} visible pages "
          f"({len(pages) - len(visible_pages)} hidden)")
    for p in visible_pages:
        n_visuals = len(p['visuals'])
        print(f"  - {_safe(p['display_name'])} ({n_visuals} visuals)")

    # Step 2: Extract model metadata
    print("\nExtracting semantic model metadata...")
    model = extract_model_metadata(str(pbip_root))
    print(f"  Found {len(model['tables'])} tables, "
          f"{len(model['measures'])} measures, "
          f"{len(model['relationships'])} relationships")

    # Step 3: Build DAX queries
    print("\nBuilding DAX queries for each page...")
    dax_queries = build_dax_queries(visible_pages, model)
    total_queries = sum(len(d['queries']) for d in dax_queries)
    print(f"  Generated {total_queries} queries across {len(dax_queries)} pages")

    # Step 4: Find images for each page
    # Priority 1: Live screenshots from the running Power BI Desktop instance
    # Pass the PBIP stem so we select the right window when multiple PBI files are open
    pbip_stem = pbip_path.stem if pbip_path.suffix.lower() == '.pbip' else pbip_root.name
    print("\nCapturing live screenshots from Power BI Desktop...")
    print("  (Close the Fields / Visualizations / Filters panels for cleanest results)")
    page_images = _capture_pbi_desktop_screenshots(visible_pages, pbip_stem=pbip_stem)

    if not page_images:
        # Priority 2: Companion PDF/PPTX in the same folder
        print("  Checking for companion PDF/PPTX export in the same folder...")
        page_images = _extract_companion_images(pbip_path, pbip_root, len(visible_pages))
        if page_images:
            matched = sum(1 for v in page_images.values() if v)
            print(f"  Found images for {matched} of {len(visible_pages)} pages from companion file")
        else:
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
