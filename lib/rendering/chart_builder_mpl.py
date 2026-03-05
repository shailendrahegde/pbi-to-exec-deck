"""
Matplotlib Chart Renderer — high-quality PNG output for PPTX embedding.

All 23 chart types (+ aliases) are rendered as 200 DPI PNG images, embedded via
slide.shapes.add_picture(). This gives full pixel-level control:
antialiasing, per-point colors, proper label placement, and chart types
that python-pptx renders coarsely (scatter, bubble, treemap, funnel, gauge).

Both renderers coexist. builder.py routes to this module when use_mpl=True.
"""

import io
import math
import matplotlib
matplotlib.use('Agg')          # headless — no display server needed
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

from lib.analysis.insights import ChartSpec

# ---------------------------------------------------------------------------
# PBI-matching color palettes (hex strings for matplotlib)
# ---------------------------------------------------------------------------

PBI_BLUE_DARK    = "#003278"
PBI_BLUE         = "#0072C6"
PBI_BLUE_MID     = "#64A5DC"
PBI_BLUE_LIGHT   = "#BDDCf3"
PBI_PURPLE_DARK  = "#5B1F6E"
PBI_PURPLE       = "#7B3F8B"
PBI_PURPLE_MID   = "#9B64B2"
PBI_PURPLE_LIGHT = "#C8AAD9"
PBI_PINK_DARK    = "#A03070"
PBI_PINK         = "#C84B97"
PBI_PINK_MID     = "#E07ABB"
PBI_PINK_LIGHT   = "#F0B9DA"
DARK_GRAY        = "#595959"
MID_GRAY         = "#8C8C8C"
LIGHT_GRAY       = "#C8C8C8"
GRID_GRAY        = "#E8E8E8"
WHITE            = "#FFFFFF"

MULTI_PAL = [
    PBI_BLUE, PBI_PINK, PBI_PURPLE, PBI_BLUE_DARK,
    PBI_PINK_DARK, PBI_PURPLE_MID, PBI_BLUE_MID,
]
HABIT_PAL = ["#C3C3C3", PBI_BLUE_LIGHT, PBI_BLUE, PBI_BLUE_DARK]


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _get_fig(w_in: float, h_in: float):
    return plt.figure(figsize=(w_in, h_in), facecolor='white', dpi=100)


def _to_png(fig, dpi: int = 200) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close(fig)
    return buf.getvalue()


def _style_ax(ax, grid='y', bottom_spine=True):
    """Clean PBI-style axes: no top/right/left spine, light gray grid."""
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_color(LIGHT_GRAY)
    ax.spines['bottom'].set_visible(bottom_spine)
    if grid == 'both':
        ax.grid(axis='x', color=GRID_GRAY, linewidth=0.5, alpha=0.8, zorder=0)
        ax.grid(axis='y', color=GRID_GRAY, linewidth=0.5, alpha=0.8, zorder=0)
    elif grid:
        ax.grid(axis=grid, color=GRID_GRAY, linewidth=0.5, alpha=0.8, zorder=0)
    ax.tick_params(colors=DARK_GRAY, labelsize=8)


def _hex_to_mpl(h: str) -> str:
    """Validate and return hex color; fall back to PBI_BLUE."""
    if h and isinstance(h, str) and h.startswith('#') and len(h) == 7:
        return h
    return PBI_BLUE


def _point_colors(data, highlight: str) -> list:
    """Return per-bar color list respecting highlight and explicit colors."""
    hl = (highlight or '').lower()
    colors = []
    for dp in data:
        if dp.color:
            colors.append(_hex_to_mpl(dp.color))
        elif hl and dp.label.lower() == hl:
            colors.append(PBI_BLUE)
        elif hl:
            colors.append(PBI_BLUE_LIGHT)
        else:
            colors.append(PBI_BLUE)
    return colors


def _render_placeholder(spec: ChartSpec, w: float, h: float, dpi: int,
                         error: str = "") -> bytes:
    """Gray box with chart type label — shown for unknown or failed types."""
    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_facecolor('#F5F5F5')
    ax.set_axis_off()
    msg = f"Chart: {spec.type or 'unknown'}"
    if error:
        msg += f"\n({error[:60]})"
    ax.text(0.5, 0.5, msg, ha='center', va='center',
            fontsize=9, color=DARK_GRAY, transform=ax.transAxes)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# Per-type renderers  (signature: fn(spec, w_in, h_in, dpi) -> bytes)
# ---------------------------------------------------------------------------

def _render_bar(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Horizontal bar chart — PBI-style per-point colouring."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    labels = [dp.label for dp in data]
    values = [dp.value for dp in data]
    colors = _point_colors(data, spec.highlight)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    y = list(range(len(labels)))
    bars = ax.barh(y, values, color=colors, height=0.6, zorder=3)
    ax.bar_label(bars, padding=4, fontsize=7, color=DARK_GRAY)
    ax.set_xlim(0, max(values) * 1.28 if values else 1)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.invert_yaxis()
    _style_ax(ax, grid='x')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_column(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Vertical column chart — PBI-style per-point colouring."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    labels = [dp.label for dp in data]
    values = [dp.value for dp in data]
    colors = _point_colors(data, spec.highlight)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    x = list(range(len(labels)))
    bars = ax.bar(x, values, color=colors, width=0.6, zorder=3)
    ax.bar_label(bars, padding=3, fontsize=7, color=DARK_GRAY)
    ax.set_ylim(0, max(values) * 1.28 if values else 1)
    ax.set_xticks(x)
    max_len = max((len(l) for l in labels), default=0)
    rot = 30 if max_len > 6 else 0
    ax.set_xticklabels(labels, fontsize=8, rotation=rot,
                        ha='right' if rot else 'center')
    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_bar_stacked(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Stacked horizontal bar chart."""
    series = spec.series
    if not series:
        return _render_bar(spec, w, h, dpi)

    # Build ordered category and series lists
    cat_order = list(dict.fromkeys(s.get('label', '') for s in series))
    ser_names = list(dict.fromkeys(s.get('series', '') for s in series))
    data_map: dict = {sn: {c: 0.0 for c in cat_order} for sn in ser_names}
    for item in series:
        data_map[item.get('series', '')][item.get('label', '')] = float(item.get('value', 0))

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    y = list(range(len(cat_order)))
    bottoms = np.zeros(len(cat_order))
    for i, sn in enumerate(ser_names):
        vals = np.array([data_map[sn][c] for c in cat_order])
        ax.barh(y, vals, left=bottoms, label=sn,
                color=MULTI_PAL[i % len(MULTI_PAL)], height=0.6, zorder=3)
        bottoms += vals

    ax.set_yticks(y)
    ax.set_yticklabels(cat_order, fontsize=8)
    ax.invert_yaxis()
    _style_ax(ax, grid='x')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    ncol = min(4, len(ser_names))
    ax.legend(loc='lower right', ncol=ncol, fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_column_stacked(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Stacked vertical column chart."""
    series = spec.series
    if not series:
        return _render_column(spec, w, h, dpi)

    cat_order = list(dict.fromkeys(s.get('label', '') for s in series))
    ser_names = list(dict.fromkeys(s.get('series', '') for s in series))
    data_map: dict = {sn: {c: 0.0 for c in cat_order} for sn in ser_names}
    for item in series:
        data_map[item.get('series', '')][item.get('label', '')] = float(item.get('value', 0))

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    x = list(range(len(cat_order)))
    bottoms = np.zeros(len(cat_order))
    for i, sn in enumerate(ser_names):
        vals = np.array([data_map[sn][c] for c in cat_order])
        ax.bar(x, vals, bottom=bottoms, label=sn,
               color=MULTI_PAL[i % len(MULTI_PAL)], width=0.6, zorder=3)
        bottoms += vals

    ax.set_xticks(x)
    max_len = max((len(c) for c in cat_order), default=0)
    rot = 30 if max_len > 6 else 0
    ax.set_xticklabels(cat_order, fontsize=8, rotation=rot,
                        ha='right' if rot else 'center')
    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    ncol = min(4, len(ser_names))
    ax.legend(loc='upper right', ncol=ncol, fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_line(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Line chart with markers — multi-series, PBI-styled."""
    series = spec.series
    if not series:
        return _render_placeholder(spec, w, h, dpi)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    xs_labels = []
    for i, s in enumerate(series):
        pts = s.get('points', [])
        xs_labels = [str(p.get('x', '')) for p in pts]
        ys = [float(p.get('y', 0)) for p in pts]
        c = s.get('color') or MULTI_PAL[i % len(MULTI_PAL)]
        ax.plot(range(len(xs_labels)), ys, color=c, linewidth=2,
                marker='o', markersize=4, label=s.get('name', ''), zorder=3)

    if xs_labels:
        ax.set_xticks(range(len(xs_labels)))
        max_len = max((len(x) for x in xs_labels), default=0)
        rot = 30 if max_len > 4 else 0
        ax.set_xticklabels(xs_labels, fontsize=7, rotation=rot,
                            ha='right' if rot else 'center')

    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    if len(series) > 1:
        ax.legend(loc='upper left', fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_area(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Area chart with fill — multi-series, PBI-styled."""
    series = spec.series
    if not series:
        return _render_placeholder(spec, w, h, dpi)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    xs_labels = []
    for i, s in enumerate(series):
        pts = s.get('points', [])
        xs_labels = [str(p.get('x', '')) for p in pts]
        ys = [float(p.get('y', 0)) for p in pts]
        c = s.get('color') or MULTI_PAL[i % len(MULTI_PAL)]
        ax.plot(range(len(xs_labels)), ys, color=c, linewidth=2,
                label=s.get('name', ''), zorder=3)
        ax.fill_between(range(len(xs_labels)), ys, alpha=0.2, color=c, zorder=2)

    if xs_labels:
        ax.set_xticks(range(len(xs_labels)))
        max_len = max((len(x) for x in xs_labels), default=0)
        rot = 30 if max_len > 4 else 0
        ax.set_xticklabels(xs_labels, fontsize=7, rotation=rot,
                            ha='right' if rot else 'center')

    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    if len(series) > 1:
        ax.legend(loc='upper left', fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_pie(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Pie chart — per-slice colors, percent labels."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    labels = [dp.label for dp in data]
    values = [dp.value for dp in data]
    colors = [
        _hex_to_mpl(dp.color) if dp.color else MULTI_PAL[i % len(MULTI_PAL)]
        for i, dp in enumerate(data)
    ]

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)

    if len(data) > 4:
        wedges, _, autotexts = ax.pie(
            values, colors=colors, autopct='%1.0f%%', pctdistance=0.75,
            startangle=90, counterclock=False,
            wedgeprops={'edgecolor': 'white', 'linewidth': 1.5}
        )
        for at in autotexts:
            at.set_fontsize(7)
            at.set_color('white')
        ax.legend(wedges, labels, loc='center left', bbox_to_anchor=(1, 0.5),
                  fontsize=7, frameon=False)
    else:
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, colors=colors, autopct='%1.0f%%',
            pctdistance=0.75, startangle=90, counterclock=False,
            wedgeprops={'edgecolor': 'white', 'linewidth': 1.5}
        )
        for t in texts:
            t.set_fontsize(7)
        for at in autotexts:
            at.set_fontsize(7)
            at.set_color('white')

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_donut(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Doughnut chart — habit-tier palette, legend on right."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    labels = [dp.label for dp in data]
    values = [dp.value for dp in data]
    default_pal = HABIT_PAL if len(data) == 4 else MULTI_PAL
    colors = [
        _hex_to_mpl(dp.color) if dp.color else default_pal[i % len(default_pal)]
        for i, dp in enumerate(data)
    ]

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    wedges, _, autotexts = ax.pie(
        values, colors=colors, autopct='%1.0f%%', pctdistance=0.75,
        startangle=90, counterclock=False,
        wedgeprops={'width': 0.55, 'edgecolor': 'white', 'linewidth': 1.5}
    )
    for at in autotexts:
        at.set_fontsize(7)
        at.set_color('white')

    ax.legend(wedges, labels, loc='center left', bbox_to_anchor=(1, 0.5),
              fontsize=7, frameon=False)
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_scatter(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """XY scatter — per-point colors and labels, axis titles."""
    series = spec.series
    if not series:
        return _render_placeholder(spec, w, h, dpi)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    for s in series:
        x = float(s.get('x', 0))
        y = float(s.get('y', 0))
        is_hl = s.get('highlight', False)
        c = s.get('color') or (PBI_PINK if is_hl else PBI_PINK_LIGHT)
        sz = 100 if is_hl else 60
        ax.scatter([x], [y], s=sz, c=c, alpha=0.85,
                   edgecolors='white', linewidth=1.5, zorder=3)
        name = s.get('name', '')
        if name:
            ax.annotate(name, (x, y), xytext=(0, 8),
                        textcoords='offset points', fontsize=7,
                        ha='center', color=DARK_GRAY)

    _style_ax(ax, grid='both')
    if spec.x_label:
        ax.set_xlabel(spec.x_label, fontsize=8, color=DARK_GRAY)
    if spec.y_label:
        ax.set_ylabel(spec.y_label, fontsize=8, color=DARK_GRAY)
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_bubble(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Bubble chart — size-normalized dots, per-point colors and labels."""
    series = spec.series
    if not series:
        return _render_placeholder(spec, w, h, dpi)

    raw_sizes = [float(s.get('size', 5)) for s in series]
    mn, mx = min(raw_sizes), max(raw_sizes)
    rng = mx - mn if mx > mn else 1.0
    norm_sizes = [100 + (v - mn) / rng * 1100 for v in raw_sizes]

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    for i, (s, sz) in enumerate(zip(series, norm_sizes)):
        x = float(s.get('x', 0))
        y = float(s.get('y', 0))
        is_hl = s.get('highlight', False)
        c = s.get('color') or (PBI_PINK if is_hl else MULTI_PAL[i % len(MULTI_PAL)])
        ax.scatter([x], [y], s=sz, c=c, alpha=0.8,
                   edgecolors='white', linewidth=1.5, zorder=3)
        name = s.get('name', '')
        if name:
            ax.annotate(name, (x, y), xytext=(0, 8),
                        textcoords='offset points', fontsize=7,
                        ha='center', color=DARK_GRAY)

    _style_ax(ax, grid='both')
    if spec.x_label:
        ax.set_xlabel(spec.x_label, fontsize=8, color=DARK_GRAY)
    if spec.y_label:
        ax.set_ylabel(spec.y_label, fontsize=8, color=DARK_GRAY)
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_radar(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Radar / spider chart — polar projection, multi-series fill."""
    series = spec.series
    if not series:
        return _render_placeholder(spec, w, h, dpi)

    pts0 = series[0].get('points', [])
    categories = [str(p.get('x', '')) for p in pts0]
    N = len(categories)
    if N < 3:
        return _render_placeholder(spec, w, h, dpi)

    angles = [2 * math.pi * i / N for i in range(N)]
    angles_closed = angles + [angles[0]]

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111, projection='polar')

    for i, s in enumerate(series):
        pts = s.get('points', [])
        vals = [float(p.get('y', 0)) for p in pts]
        if len(vals) < N:
            vals.extend([0] * (N - len(vals)))
        vals_closed = vals + [vals[0]]
        c = s.get('color') or MULTI_PAL[i % len(MULTI_PAL)]
        ax.plot(angles_closed, vals_closed, color=c, linewidth=2,
                label=s.get('name', ''))
        ax.fill(angles_closed, vals_closed, color=c, alpha=0.15)

    ax.set_xticks(angles)
    ax.set_xticklabels(categories, fontsize=7, color=DARK_GRAY)
    ax.tick_params(colors=DARK_GRAY, labelsize=7)
    ax.set_title(spec.title or '', pad=14, fontsize=9, color=DARK_GRAY)
    if len(series) > 1:
        ax.legend(loc='lower right', fontsize=7, frameon=False)

    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_kpi(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Single KPI card — large bold value, label, optional subtitle."""
    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_facecolor('#F8FAFD')
    ax.set_axis_off()

    ax.text(0.5, 0.62, spec.value or '', fontsize=28, fontweight='bold',
            color=PBI_BLUE, ha='center', va='center', transform=ax.transAxes)
    ax.text(0.5, 0.38, spec.label or '', fontsize=10,
            color=DARK_GRAY, ha='center', va='center', transform=ax.transAxes)
    if spec.subtitle:
        ax.text(0.5, 0.22, spec.subtitle, fontsize=8,
                color=MID_GRAY, ha='center', va='center', transform=ax.transAxes)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


def _render_kpi_row(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Horizontal row of 2-4 compact KPI tiles."""
    kpis = spec.series or []
    if not kpis:
        return _render_kpi(spec, w, h, dpi)

    n = len(kpis)
    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    tile_w = 1.0 / n
    gap = 0.01
    val_fs  = max(10, min(22, int(22 / math.sqrt(n))))
    lbl_fs  = max(7,  min(10, int(10 / math.sqrt(n))))
    sub_fs  = max(6,  min(8,  int(8  / math.sqrt(n))))

    for i, kpi in enumerate(kpis):
        x0 = i * tile_w + gap / 2
        tw = tile_w - gap
        cx = x0 + tw / 2

        # Tile background
        rect = mpatches.FancyBboxPatch(
            (x0, 0.05), tw, 0.90,
            boxstyle="round,pad=0.01",
            facecolor='#F8FAFD',
            edgecolor=LIGHT_GRAY, linewidth=0.5,
            transform=ax.transAxes, clip_on=False
        )
        ax.add_patch(rect)

        ax.text(cx, 0.65, kpi.get('value', ''), fontsize=val_fs,
                fontweight='bold', color=PBI_BLUE, ha='center', va='center',
                transform=ax.transAxes)
        ax.text(cx, 0.38, kpi.get('label', ''), fontsize=lbl_fs,
                color=DARK_GRAY, ha='center', va='center',
                transform=ax.transAxes)
        if kpi.get('subtitle'):
            ax.text(cx, 0.20, kpi['subtitle'], fontsize=sub_fs,
                    color=MID_GRAY, ha='center', va='center',
                    transform=ax.transAxes)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


def _render_heatmap(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Heatmap with Blues colormap and cell text labels."""
    rows = spec.rows or []
    cols = spec.columns or []
    vals = spec.values or []

    if not rows or not cols or not vals:
        return _render_placeholder(spec, w, h, dpi)

    data = np.array(
        [[vals[ri][ci] if ci < len(vals[ri]) else 0.0
          for ci in range(len(cols))]
         for ri in range(len(rows))],
        dtype=float
    )

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    im = ax.imshow(data, cmap='Blues', aspect='auto',
                   vmin=0, vmax=data.max() if data.max() > 0 else 1)

    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(cols, fontsize=7, rotation=30, ha='right')
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels(rows, fontsize=7)

    norm_max = data.max() if data.max() > 0 else 1.0
    cell_fs = max(6, min(9, int(70 / max(len(rows), len(cols)))))
    fmt = spec.value_format or '.1f'
    for ri in range(len(rows)):
        for ci in range(len(cols)):
            val = data[ri, ci]
            norm = val / norm_max
            txt_color = 'white' if norm > 0.6 else DARK_GRAY
            try:
                lbl = format(val, fmt)
            except Exception:
                lbl = str(val)
            ax.text(ci, ri, lbl, ha='center', va='center',
                    fontsize=cell_fs, color=txt_color)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    for sp in ax.spines.values():
        sp.set_visible(False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_table(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Table with dark header, alternating rows, highlighted column."""
    cols = spec.table_columns or []
    rows = spec.table_rows or []
    if not cols or not rows:
        return _render_placeholder(spec, w, h, dpi)

    n_rows = len(rows)
    fs = max(6, min(9, 56 // max(n_rows + 1, 1)))

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_axis_off()

    tbl = ax.table(
        cellText=rows,
        colLabels=cols,
        loc='center',
        cellLoc='center'
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(fs)
    tbl.scale(1, max(1.5, 1.8 / max(n_rows + 1, 1) * (h / w * 4)))

    # Header: dark blue background, white bold text
    for ci in range(len(cols)):
        cell = tbl[0, ci]
        cell.set_facecolor(PBI_BLUE_DARK)
        cell.get_text().set_color('white')
        cell.get_text().set_fontweight('bold')
        cell.get_text().set_fontsize(fs)

    # Data rows: alternating backgrounds, bold highlight column
    alt_bg = '#F2F6FC'
    for ri in range(n_rows):
        bg = alt_bg if ri % 2 == 0 else WHITE
        for ci in range(len(cols)):
            cell = tbl[ri + 1, ci]
            cell.set_facecolor(bg)
            cell.get_text().set_fontsize(fs)
            if ci == spec.highlight_col:
                cell.get_text().set_fontweight('bold')
                cell.get_text().set_color(PBI_BLUE_DARK)
            else:
                cell.get_text().set_color(DARK_GRAY)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


def _render_funnel(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Funnel chart — centered bars that narrow per stage value."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    n = len(data)
    max_val = max(dp.value for dp in data) if data else 1.0
    palette = [PBI_BLUE_DARK, PBI_BLUE, PBI_BLUE_MID, PBI_BLUE_LIGHT, LIGHT_GRAY]

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    row_h = 0.88 / n
    for i, dp in enumerate(data):
        frac = dp.value / max_val if max_val > 0 else 0
        bar_w = frac * 0.60
        x0 = 0.5 - bar_w / 2
        y0 = 0.94 - (i + 1) * row_h
        pad = 0.008

        c = _hex_to_mpl(dp.color) if dp.color else palette[i % len(palette)]
        rect = mpatches.FancyBboxPatch(
            (x0, y0 + pad), bar_w, row_h - 2 * pad,
            boxstyle="round,pad=0.003",
            facecolor=c, edgecolor='white', linewidth=0.5,
            transform=ax.transAxes
        )
        ax.add_patch(rect)

        # Label to the left
        ax.text(max(0, x0 - 0.02), y0 + row_h / 2, dp.label,
                ha='right', va='center', fontsize=8, color=DARK_GRAY,
                transform=ax.transAxes)

        # Value to the right
        try:
            vt = format(dp.value, '.0f') if dp.value == int(dp.value) else format(dp.value, '.1f')
        except Exception:
            vt = str(dp.value)
        ax.text(min(1, x0 + bar_w + 0.02), y0 + row_h / 2, vt,
                ha='left', va='center', fontsize=8, fontweight='bold',
                color=DARK_GRAY, transform=ax.transAxes)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


def _squarify_mpl(items):
    """
    Pure-Python squarified treemap layout.
    Returns: list of (label, value, x0, y0, w, h) in normalized [0,1] coords.
    """
    if not items:
        return []
    total = sum(v for _, v in items)
    if total <= 0:
        return []
    result = []

    def layout_row(row, x0, y0, rw, rh, horiz):
        rs = sum(v for _, v in row)
        if rs <= 0:
            return
        off = 0.0
        for lbl, val in row:
            f = val / rs
            if horiz:
                result.append((lbl, val, x0, y0 + off, rw, rh * f))
                off += rh * f
            else:
                result.append((lbl, val, x0 + off, y0, rw * f, rh))
                off += rw * f

    def recurse(its, x0, y0, w0, h0):
        if not its:
            return
        tot = sum(v for _, v in its)
        if tot <= 0:
            return
        horiz = w0 <= h0
        cur, best = [], float('inf')
        for item in its:
            trial = cur + [item]
            rs = sum(v for _, v in trial)
            frac = rs / tot
            if horiz:
                rh = h0 * frac
                ratios = [(v / rs * rh / w0 if w0 > 0 else 1e9) for _, v in trial]
            else:
                rw = w0 * frac
                ratios = [(v / rs * rw / h0 if h0 > 0 else 1e9) for _, v in trial]
            mr = max((max(r, 1 / r) if r > 0 else 1e9) for r in ratios)
            if mr <= best:
                best, cur = mr, trial
            else:
                break
        if not cur:
            cur = [its[0]]
        rs = sum(v for _, v in cur)
        frac = rs / tot
        if horiz:
            rh = h0 * frac
            layout_row(cur, x0, y0, w0, rh, True)
            recurse([it for it in its if it not in cur], x0, y0 + rh, w0, h0 - rh)
        else:
            rw = w0 * frac
            layout_row(cur, x0, y0, rw, h0, False)
            recurse([it for it in its if it not in cur], x0 + rw, y0, w0 - rw, h0)

    recurse(sorted(items, key=lambda x: x[1], reverse=True), 0.0, 0.0, 1.0, 1.0)
    return result


def _render_treemap(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """Treemap with squarified layout and PBI palette."""
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    items = [(dp.label, dp.value) for dp in data]
    rects = _squarify_mpl(items)
    dmap = {dp.label: dp for dp in data}

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    ax.set_axis_off()
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    gap = 0.005
    for i, (lbl, val, x0, y0, rw, rh) in enumerate(rects):
        dp = dmap.get(lbl)
        c = _hex_to_mpl(dp.color) if (dp and dp.color) else MULTI_PAL[i % len(MULTI_PAL)]
        rect = mpatches.FancyBboxPatch(
            (x0 + gap, y0 + gap),
            max(rw - 2 * gap, 0.002),
            max(rh - 2 * gap, 0.002),
            boxstyle="square,pad=0",
            facecolor=c, edgecolor='white', linewidth=1.0,
            transform=ax.transAxes
        )
        ax.add_patch(rect)

        # Font size scales with cell area
        area = rw * rh
        fs = max(6, min(11, int(math.sqrt(area) * 28)))
        if fs >= 6 and rw > 0.06 and rh > 0.06:
            ax.text(x0 + gap + rw * 0.07, y0 + gap + rh * 0.55, lbl,
                    fontsize=fs, color='white', va='center',
                    ha='left', transform=ax.transAxes, clip_on=True)

    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


def _render_gauge(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """
    Half-donut gauge chart.
    With startangle=180 and counterclock=False, the filled+empty arcs
    occupy the top semicircle; the hidden (white) slice sits below.
    """
    try:
        raw = float(spec.value or 0)
    except (ValueError, TypeError):
        raw = 0.0
    max_val = spec.max_value or 100.0
    frac = max(0.0, min(1.0, raw / max_val))
    filled = frac * 50.0
    empty  = (1 - frac) * 50.0
    hidden = 50.0

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)

    ax.pie(
        [filled, empty, hidden],
        colors=[PBI_BLUE, LIGHT_GRAY, 'white'],
        startangle=180, counterclock=False,
        wedgeprops={'width': 0.45, 'edgecolor': 'white', 'linewidth': 1.5}
    )

    # Show top half only; leave room below for text
    ax.set_ylim(-0.40, 1)

    # Value text just below center
    val_txt = f"{int(round(raw))}{'%' if max_val == 100 else ''}"
    ax.text(0, -0.08, val_txt, fontsize=22, fontweight='bold',
            color=PBI_BLUE, ha='center', va='center')
    if spec.label:
        ax.text(0, -0.30, spec.label, fontsize=9, color=DARK_GRAY,
                ha='center', va='center')

    if spec.title:
        ax.set_title(spec.title, loc='left', fontsize=9, color=DARK_GRAY, pad=4)

    fig.tight_layout(pad=0.3)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# Waterfall chart
# ---------------------------------------------------------------------------

def _render_waterfall(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """
    Waterfall chart — shows incremental positive/negative changes.

    Positive values are increases (green), negative are decreases (red).
    Items with ``"color"`` set or whose label contains "total" (case-insensitive)
    are rendered as total bars anchored at zero.
    """
    data = spec.data or []
    if not data:
        return _render_placeholder(spec, w, h, dpi)

    labels = [dp.label for dp in data]
    values = [dp.value for dp in data]
    n = len(data)

    # Determine which bars are "total" bars (anchored at zero)
    is_total = []
    for dp in data:
        if dp.color:
            is_total.append(True)
        elif 'total' in dp.label.lower() or 'net' in dp.label.lower():
            is_total.append(True)
        else:
            is_total.append(False)

    # Compute running baseline for each bar
    bottoms = []
    running = 0.0
    for i, (val, tot) in enumerate(zip(values, is_total)):
        if tot:
            bottoms.append(0.0)
            running = val  # reset running total to this bar's value
        else:
            if val >= 0:
                bottoms.append(running)
                running += val
            else:
                running += val
                bottoms.append(running)

    INC_COLOR = "#2E7D32"   # green for increases
    DEC_COLOR = "#C62828"   # red for decreases
    TOT_COLOR = PBI_BLUE_DARK  # blue for totals

    colors = []
    bar_values = []
    for i, (val, tot) in enumerate(zip(values, is_total)):
        if tot:
            colors.append(TOT_COLOR)
            bar_values.append(val)
        elif val >= 0:
            colors.append(INC_COLOR)
            bar_values.append(val)
        else:
            colors.append(DEC_COLOR)
            bar_values.append(abs(val))

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    x = list(range(n))
    bars = ax.bar(x, bar_values, bottom=bottoms, color=colors, width=0.6,
                  zorder=3, edgecolor='white', linewidth=0.5)

    # Data labels
    for i, (bar, val, tot) in enumerate(zip(bars, values, is_total)):
        y_pos = bar.get_y() + bar.get_height() + (max(abs(v) for v in values) * 0.02)
        prefix = "+" if val > 0 and not tot else ""
        try:
            txt = f"{prefix}{val:.0f}" if val == int(val) else f"{prefix}{val:.1f}"
        except (ValueError, TypeError):
            txt = str(val)
        ax.text(i, y_pos, txt, ha='center', va='bottom', fontsize=7,
                color=DARK_GRAY, fontweight='bold' if tot else 'normal')

    # Connector lines between bars (skip to/from totals)
    for i in range(n - 1):
        if is_total[i] or is_total[i + 1]:
            continue
        y_conn = bottoms[i] + (values[i] if values[i] >= 0 else 0)
        ax.plot([i + 0.3, i + 0.7], [y_conn, y_conn],
                color=LIGHT_GRAY, linewidth=0.8, zorder=2)

    ax.set_xticks(x)
    max_len = max((len(l) for l in labels), default=0)
    rot = 30 if max_len > 6 else 0
    ax.set_xticklabels(labels, fontsize=8, rotation=rot,
                        ha='right' if rot else 'center')
    _style_ax(ax, grid='y')
    ax.axhline(0, color=LIGHT_GRAY, linewidth=0.8, zorder=1)
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# Combo chart (column + line)
# ---------------------------------------------------------------------------

def _render_combo(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """
    Combo chart — column bars from ``spec.data`` with a line overlay
    from ``spec.series``.  Dual Y-axes: left for bars, right for line.
    If only data is provided, falls back to a plain column chart.
    If only series is provided, falls back to a plain line chart.
    """
    has_data = bool(spec.data)
    has_series = bool(spec.series)

    if not has_data and not has_series:
        return _render_placeholder(spec, w, h, dpi)
    if not has_series:
        return _render_column(spec, w, h, dpi)
    if not has_data:
        return _render_line(spec, w, h, dpi)

    labels = [dp.label for dp in spec.data]
    values = [dp.value for dp in spec.data]
    colors = _point_colors(spec.data, spec.highlight)
    n = len(labels)

    fig = _get_fig(w, h)
    ax1 = fig.add_subplot(111)
    x = list(range(n))
    ax1.bar(x, values, color=colors, width=0.6, zorder=3, alpha=0.85)
    ax1.set_ylabel(spec.y_label or '', fontsize=8, color=DARK_GRAY)
    ax1.set_ylim(0, max(values) * 1.25 if values else 1)

    # Line overlay on secondary axis
    ax2 = ax1.twinx()
    for i, s in enumerate(spec.series):
        pts = s.get('points', [])
        line_xs = []
        line_ys = []
        for p in pts:
            # Try to map x-labels to bar positions
            px = str(p.get('x', ''))
            if px in labels:
                line_xs.append(labels.index(px))
            else:
                try:
                    line_xs.append(int(p.get('x', i)))
                except (ValueError, TypeError):
                    line_xs.append(len(line_xs))
            line_ys.append(float(p.get('y', 0)))
        c = s.get('color') or MULTI_PAL[(i + 3) % len(MULTI_PAL)]
        ax2.plot(line_xs, line_ys, color=c, linewidth=2, marker='o',
                 markersize=5, label=s.get('name', ''), zorder=5)
        # Data labels on line
        for lx, ly in zip(line_xs, line_ys):
            try:
                txt = f"{ly:.0f}" if ly == int(ly) else f"{ly:.1f}"
            except (ValueError, TypeError):
                txt = str(ly)
            ax2.annotate(txt, (lx, ly), xytext=(0, 8),
                         textcoords='offset points', fontsize=6.5,
                         ha='center', color=c, fontweight='bold')

    ax2.spines['top'].set_visible(False)
    ax2.spines['left'].set_visible(False)
    ax2.spines['right'].set_color(LIGHT_GRAY)
    ax2.tick_params(colors=DARK_GRAY, labelsize=7)
    if spec.x_label:
        ax2.set_ylabel(spec.x_label, fontsize=8, color=DARK_GRAY)  # x_label used for right-axis label

    ax1.set_xticks(x)
    max_len = max((len(l) for l in labels), default=0)
    rot = 30 if max_len > 6 else 0
    ax1.set_xticklabels(labels, fontsize=8, rotation=rot,
                         ha='right' if rot else 'center')
    _style_ax(ax1, grid='y')

    # Combined legend
    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    if handles2:
        ax2.legend(handles1 + handles2, labels1 + labels2,
                   loc='upper left', fontsize=7, frameon=False)

    ax1.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# 100% Stacked Bar / Column
# ---------------------------------------------------------------------------

def _render_bar_stacked_100(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """100% stacked horizontal bar — each category sums to 100%."""
    series = spec.series
    if not series:
        return _render_bar(spec, w, h, dpi)

    cat_order = list(dict.fromkeys(s.get('label', '') for s in series))
    ser_names = list(dict.fromkeys(s.get('series', '') for s in series))
    data_map: dict = {sn: {c: 0.0 for c in cat_order} for sn in ser_names}
    for item in series:
        data_map[item.get('series', '')][item.get('label', '')] = float(item.get('value', 0))

    # Compute totals for normalization
    totals = {c: sum(data_map[sn][c] for sn in ser_names) for c in cat_order}

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    y = list(range(len(cat_order)))
    bottoms = np.zeros(len(cat_order))
    for i, sn in enumerate(ser_names):
        vals_raw = np.array([data_map[sn][c] for c in cat_order])
        tots = np.array([totals[c] if totals[c] > 0 else 1.0 for c in cat_order])
        vals = vals_raw / tots * 100.0
        ax.barh(y, vals, left=bottoms, label=sn,
                color=MULTI_PAL[i % len(MULTI_PAL)], height=0.6, zorder=3)
        # Center-label percentages ≥ 8%
        for j, v in enumerate(vals):
            if v >= 8:
                ax.text(bottoms[j] + v / 2, j, f"{v:.0f}%",
                        ha='center', va='center', fontsize=6.5, color='white',
                        fontweight='bold')
        bottoms += vals

    ax.set_xlim(0, 100)
    ax.set_yticks(y)
    ax.set_yticklabels(cat_order, fontsize=8)
    ax.invert_yaxis()
    _style_ax(ax, grid='x')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    ncol = min(4, len(ser_names))
    ax.legend(loc='lower right', ncol=ncol, fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


def _render_column_stacked_100(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """100% stacked vertical column — each category sums to 100%."""
    series = spec.series
    if not series:
        return _render_column(spec, w, h, dpi)

    cat_order = list(dict.fromkeys(s.get('label', '') for s in series))
    ser_names = list(dict.fromkeys(s.get('series', '') for s in series))
    data_map: dict = {sn: {c: 0.0 for c in cat_order} for sn in ser_names}
    for item in series:
        data_map[item.get('series', '')][item.get('label', '')] = float(item.get('value', 0))

    totals = {c: sum(data_map[sn][c] for sn in ser_names) for c in cat_order}

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    x = list(range(len(cat_order)))
    bottoms = np.zeros(len(cat_order))
    for i, sn in enumerate(ser_names):
        vals_raw = np.array([data_map[sn][c] for c in cat_order])
        tots = np.array([totals[c] if totals[c] > 0 else 1.0 for c in cat_order])
        vals = vals_raw / tots * 100.0
        ax.bar(x, vals, bottom=bottoms, label=sn,
               color=MULTI_PAL[i % len(MULTI_PAL)], width=0.6, zorder=3)
        for j, v in enumerate(vals):
            if v >= 8:
                ax.text(j, bottoms[j] + v / 2, f"{v:.0f}%",
                        ha='center', va='center', fontsize=6.5, color='white',
                        fontweight='bold')
        bottoms += vals

    ax.set_ylim(0, 100)
    ax.set_xticks(x)
    max_len = max((len(c) for c in cat_order), default=0)
    rot = 30 if max_len > 6 else 0
    ax.set_xticklabels(cat_order, fontsize=8, rotation=rot,
                        ha='right' if rot else 'center')
    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    ncol = min(4, len(ser_names))
    ax.legend(loc='upper right', ncol=ncol, fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# Ribbon chart
# ---------------------------------------------------------------------------

def _render_ribbon(spec: ChartSpec, w: float, h: float, dpi: int) -> bytes:
    """
    Ribbon chart — stacked area where series order can change at each x-tick,
    showing rank changes across categories.  Uses same series format as
    stacked charts (label, series, value).
    """
    series_raw = spec.series
    if not series_raw:
        return _render_placeholder(spec, w, h, dpi)

    # Parse into categories and series
    cat_order = list(dict.fromkeys(s.get('label', '') for s in series_raw))
    ser_names = list(dict.fromkeys(s.get('series', '') for s in series_raw))
    data_map: dict = {sn: {c: 0.0 for c in cat_order} for sn in ser_names}
    for item in series_raw:
        data_map[item.get('series', '')][item.get('label', '')] = float(item.get('value', 0))

    n_cats = len(cat_order)
    if n_cats < 2:
        return _render_placeholder(spec, w, h, dpi)

    fig = _get_fig(w, h)
    ax = fig.add_subplot(111)
    x = np.arange(n_cats, dtype=float)

    # At each x-tick, sort series by value (descending) and stack bottom-up
    # This makes the "ribbons" cross when ranks change
    bottoms = np.zeros(n_cats)
    drawn = {sn: {'y0': np.zeros(n_cats), 'y1': np.zeros(n_cats)} for sn in ser_names}

    for ci in range(n_cats):
        cat = cat_order[ci]
        ranked = sorted(ser_names, key=lambda sn: data_map[sn][cat], reverse=True)
        b = 0.0
        for sn in ranked:
            val = data_map[sn][cat]
            drawn[sn]['y0'][ci] = b
            drawn[sn]['y1'][ci] = b + val
            b += val

    for i, sn in enumerate(ser_names):
        c = MULTI_PAL[i % len(MULTI_PAL)]
        ax.fill_between(x, drawn[sn]['y0'], drawn[sn]['y1'],
                        color=c, alpha=0.7, label=sn, zorder=3)
        # Mid-line for visual clarity
        mid = (drawn[sn]['y0'] + drawn[sn]['y1']) / 2
        ax.plot(x, mid, color=c, linewidth=1.0, alpha=0.5, zorder=4)

    ax.set_xticks(x)
    max_len = max((len(c) for c in cat_order), default=0)
    rot = 30 if max_len > 6 else 0
    ax.set_xticklabels(cat_order, fontsize=8, rotation=rot,
                        ha='right' if rot else 'center')
    _style_ax(ax, grid='y')
    ax.set_title(spec.title or '', loc='left', fontsize=9, color=DARK_GRAY, pad=4)
    ncol = min(4, len(ser_names))
    ax.legend(loc='upper right', ncol=ncol, fontsize=7, frameon=False)
    fig.tight_layout(pad=0.4)
    return _to_png(fig, dpi)


# ---------------------------------------------------------------------------
# Dispatcher
# ---------------------------------------------------------------------------

_DISPATCH = {
    'bar':               _render_bar,
    'bar_stacked':       _render_bar_stacked,
    'bar_stacked_100':   _render_bar_stacked_100,
    'column':            _render_column,
    'column_stacked':    _render_column_stacked,
    'column_stacked_100':_render_column_stacked_100,
    'line':              _render_line,
    'area':              _render_area,
    'pie':               _render_pie,
    'donut':             _render_donut,
    'scatter':           _render_scatter,
    'bubble':            _render_bubble,
    'radar':             _render_radar,
    'kpi':               _render_kpi,
    'kpi_row':           _render_kpi_row,
    'heatmap':           _render_heatmap,
    'table':             _render_table,
    'funnel':            _render_funnel,
    'treemap':           _render_treemap,
    'gauge':             _render_gauge,
    'waterfall':         _render_waterfall,
    'combo':             _render_combo,
    'column_line':       _render_combo,        # alias
    'ribbon':            _render_ribbon,
    # PBI-native name aliases
    'card':              _render_kpi,
    'multi_row_card':    _render_kpi_row,
}


def render_chart_to_png(spec: ChartSpec, width_in: float, height_in: float,
                         dpi: int = 200) -> bytes:
    """
    Render a ChartSpec as PNG bytes at the requested size.

    Args:
        spec:       ChartSpec dataclass with chart type and data
        width_in:   Chart width  in inches (caller converts from EMU: emu / 914400)
        height_in:  Chart height in inches
        dpi:        Output resolution (default 200 DPI)

    Returns:
        PNG image as raw bytes, ready for:
            slide.shapes.add_picture(io.BytesIO(png), left, top, width, height)
    """
    fn = _DISPATCH.get((spec.type or '').lower().strip())
    if fn is None:
        return _render_placeholder(spec, width_in, height_in, dpi)
    try:
        return fn(spec, width_in, height_in, dpi)
    except Exception as e:
        print(f"  WARNING: mpl chart render failed ({spec.type}): {e}")
        return _render_placeholder(spec, width_in, height_in, dpi, error=str(e))
