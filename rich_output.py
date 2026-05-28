"""
rich_output.py — Smart return formatting for Excel MCP tools.

Three tiers based on result size:

  Tier 1 — INLINE  (≤ ROW_INLINE rows)
    → Markdown table returned as string
    → LLM renders in chat directly

  Tier 2 — HTML  (ROW_INLINE < rows ≤ ROW_HTML)
    → Self-contained HTML file with sortable/filterable table
    → Base64 PNG thumbnail of first 20 rows embedded in response
    → LLM embeds thumbnail, provides HTML file path for download

  Tier 3 — DOWNLOAD  (> ROW_HTML rows OR explicit large output)
    → Temp Excel file written to output_dir
    → Base64 PNG chart summary embedded in response
    → LLM shows chart inline, provides Excel download path
    → Raw rows NEVER passed to LLM context

Charts always → base64 PNG embedded in response string.
"""

import io, base64, json, textwrap
from pathlib import Path
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from config import MAX_INLINE_ROWS, CHART_DPI

# ── Thresholds ────────────────────────────────────────────────────────────────
ROW_INLINE  = MAX_INLINE_ROWS   # ≤ MAX_INLINE_ROWS rows → markdown table in chat
ROW_HTML    = 500               # ≤ 500 rows → HTML file + thumbnail
# > 500 rows  → Excel download + chart summary

# ── Base64 helpers ────────────────────────────────────────────────────────────

def _fig_to_base64(fig=None) -> str:
    """Convert current plt figure (or passed fig) to base64 PNG string."""
    buf = io.BytesIO()
    f = fig or plt.gcf()
    f.savefig(buf, format="png", dpi=CHART_DPI, bbox_inches="tight")
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode()
    plt.close("all")
    return b64


def _df_to_thumbnail_b64(df: pd.DataFrame, max_rows=20, title="Preview") -> str:
    """Render first max_rows of a DataFrame as a matplotlib table → base64 PNG."""
    preview = df.head(max_rows)
    ncols = len(preview.columns)
    nrows = len(preview)

    fig_w = max(8, min(ncols * 1.8, 20))
    fig_h = max(2, min(nrows * 0.4 + 1.2, 12))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    # Truncate long strings
    disp = preview.copy()
    for col in disp.select_dtypes(include="object").columns:
        disp[col] = disp[col].astype(str).str[:25]

    tbl = ax.table(
        cellText=disp.values,
        colLabels=list(disp.columns),
        cellLoc="left",
        loc="center",
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8)
    tbl.auto_set_column_width(col=list(range(ncols)))

    # Style header
    for ci in range(ncols):
        tbl[(0, ci)].set_facecolor("#1F3864")
        tbl[(0, ci)].set_text_props(color="white", fontweight="bold")

    # Alternate row colours
    for ri in range(1, nrows + 1):
        for ci in range(ncols):
            tbl[(ri, ci)].set_facecolor("#DAE3F3" if ri % 2 == 0 else "white")

    note = f"  Showing {nrows} of {len(df)} rows" if len(df) > max_rows else ""
    fig.suptitle(f"{title}{note}", fontsize=9, color="#333333", y=0.98)
    plt.tight_layout()
    b64 = _fig_to_base64(fig)
    return b64


# ── HTML table builder ────────────────────────────────────────────────────────

def _df_to_html_file(df: pd.DataFrame, title: str, out_path: Path,
                     extra_notes: str = "") -> str:
    """
    Write a self-contained, sortable, searchable HTML file.
    Returns the file path string.
    """
    # Format numbers nicely
    display_df = df.copy()
    for col in display_df.select_dtypes(include="number").columns:
        if display_df[col].abs().max() > 10000:
            display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "")
        elif display_df[col].max() <= 1.0 and display_df[col].min() >= 0:
            display_df[col] = display_df[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "")

    rows_html = ""
    for _, row in display_df.iterrows():
        cells = "".join(f"<td>{v}</td>" for v in row)
        rows_html += f"<tr>{cells}</tr>\n"

    headers_html = "".join(f"<th onclick=\"sortTable({i})\">{c} ↕</th>"
                           for i, c in enumerate(display_df.columns))

    html = textwrap.dedent(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
  body {{ font-family: 'Segoe UI', sans-serif; margin: 20px; background: #f5f7fa; color: #1a1a2e; }}
  h1 {{ color: #1F3864; font-size: 1.4em; margin-bottom: 4px; }}
  .meta {{ color: #666; font-size: 0.85em; margin-bottom: 16px; }}
  .search-box {{ margin-bottom: 12px; }}
  input#searchInput {{ padding: 8px 12px; width: 320px; border: 1px solid #ccc;
    border-radius: 6px; font-size: 0.9em; }}
  table {{ border-collapse: collapse; width: 100%; background: white;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-radius: 8px; overflow: hidden; }}
  th {{ background: #1F3864; color: white; padding: 10px 12px; text-align: left;
    cursor: pointer; font-size: 0.85em; white-space: nowrap; }}
  th:hover {{ background: #2E5090; }}
  td {{ padding: 7px 12px; font-size: 0.83em; border-bottom: 1px solid #eee; }}
  tr:nth-child(even) td {{ background: #DAE3F3; }}
  tr:hover td {{ background: #fff3cd; }}
  .count {{ color: #1F3864; font-weight: bold; font-size: 0.88em; margin-top: 8px; }}
  .notes {{ background: #e8f4fd; border-left: 4px solid #4472C4; padding: 8px 14px;
    margin-bottom: 14px; font-size: 0.85em; border-radius: 0 4px 4px 0; }}
</style>
</head>
<body>
<h1>{title}</h1>
<p class="meta">Generated by Excel MCP · {len(df):,} rows × {len(df.columns)} columns</p>
{'<div class="notes">' + extra_notes + '</div>' if extra_notes else ''}
<div class="search-box">
  <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="Search any column...">
</div>
<table id="dataTable">
  <thead><tr>{headers_html}</tr></thead>
  <tbody id="tableBody">{rows_html}</tbody>
</table>
<p class="count" id="rowCount">{len(df):,} rows displayed</p>
<script>
function filterTable() {{
  var inp = document.getElementById("searchInput").value.toLowerCase();
  var rows = document.getElementById("tableBody").getElementsByTagName("tr");
  var vis = 0;
  for (var r of rows) {{
    var txt = r.textContent.toLowerCase();
    var show = txt.includes(inp);
    r.style.display = show ? "" : "none";
    if (show) vis++;
  }}
  document.getElementById("rowCount").textContent = vis + " rows displayed";
}}
var sortDir = {{}};
function sortTable(col) {{
  var tbl = document.getElementById("tableBody");
  var rows = Array.from(tbl.getElementsByTagName("tr"));
  var asc = sortDir[col] = !sortDir[col];
  rows.sort((a,b) => {{
    var av = a.cells[col].textContent.trim();
    var bv = b.cells[col].textContent.trim();
    var an = parseFloat(av.replace(/[,%₹]/g,"")), bn = parseFloat(bv.replace(/[,%₹]/g,""));
    if (!isNaN(an) && !isNaN(bn)) return asc ? an-bn : bn-an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  }});
  rows.forEach(r => tbl.appendChild(r));
}}
</script>
</body></html>""")

    out_path.write_text(html, encoding="utf-8")
    return str(out_path)


# ── Main smart return function ────────────────────────────────────────────────

def smart_return(
    df: pd.DataFrame,
    title: str,
    output_dir: Path,
    chart_fn=None,        # optional: callable that draws a plt chart, returns nothing
    extra_notes: str = "",
    force_tier: int = 0,  # 0=auto, 1=inline, 2=html, 3=download
) -> dict:
    """
    Given a DataFrame result, decide how to return it based on size.
    Returns a dict:
      {
        "tier": 1|2|3,
        "response_text": str,          # markdown to return in MCP response
        "chart_b64": str | None,       # base64 PNG if chart exists
        "thumbnail_b64": str | None,   # base64 PNG table preview (tier 2+)
        "html_path": str | None,       # path to HTML file (tier 2)
        "excel_path": str | None,      # path to temp Excel (tier 3)
        "row_count": int,
      }
    """
    n = len(df)
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)

    # Determine tier
    tier = force_tier or (1 if n <= ROW_INLINE else 2 if n <= ROW_HTML else 3)

    result = {
        "tier": tier,
        "response_text": "",
        "chart_b64": None,
        "thumbnail_b64": None,
        "html_path": None,
        "excel_path": None,
        "row_count": n,
    }

    # ── Chart (always generate if chart_fn provided) ──────────────────────────
    if chart_fn:
        try:
            plt.close("all")
            chart_fn()
            result["chart_b64"] = _fig_to_base64()
        except Exception as e:
            result["chart_b64"] = None

    # ── Tier 1: inline markdown table ─────────────────────────────────────────
    if tier == 1:
        try:
            md_table = df.to_markdown(index=False)
        except Exception:
            md_table = df.to_string(index=False)

        lines = [f"### {title}", f"*{n} rows*", "", md_table]
        if extra_notes: lines.insert(2, f"> {extra_notes}")
        result["response_text"] = "\n".join(lines)

    # ── Tier 2: HTML file + thumbnail ─────────────────────────────────────────
    elif tier == 2:
        fname = title.lower().replace(" ", "_").replace("/", "_")[:40]
        html_path = output_dir / f"{fname}.html"
        _df_to_html_file(df, title, html_path, extra_notes)
        result["html_path"] = str(html_path)

        # Thumbnail of first 20 rows
        result["thumbnail_b64"] = _df_to_thumbnail_b64(df, max_rows=20, title=title)

        lines = [
            f"### {title}",
            f"*{n:,} rows × {len(df.columns)} columns*",
            "",
            f"📄 **HTML view (sortable/searchable):** `{html_path}`",
            "",
            "**Preview (first 20 rows):**",
        ]
        if extra_notes: lines.insert(2, f"> {extra_notes}")
        result["response_text"] = "\n".join(lines)

    # ── Tier 3: Excel download + chart summary ────────────────────────────────
    else:
        fname = title.lower().replace(" ", "_").replace("/", "_")[:40]
        excel_path = output_dir / f"{fname}_download.xlsx"

        with pd.ExcelWriter(str(excel_path), engine="openpyxl") as w:
            df.to_excel(w, sheet_name=title[:31], index=False)
        result["excel_path"] = str(excel_path)

        # Summary stats for the response
        num_cols = df.select_dtypes(include="number").columns.tolist()
        stats_lines = []
        for col in num_cols[:5]:
            s = df[col].dropna()
            if len(s):
                stats_lines.append(
                    f"  - **{col}**: total={s.sum():,.0f} | mean={s.mean():,.1f} | "
                    f"min={s.min():,.0f} | max={s.max():,.0f}"
                )

        # Thumbnail for first 20 rows
        result["thumbnail_b64"] = _df_to_thumbnail_b64(df, max_rows=20, title=f"{title} (preview)")

        lines = [
            f"### {title}",
            f"*{n:,} rows × {len(df.columns)} columns — too large for inline display*",
            "",
            f"📥 **Download Excel:** `{excel_path}`",
            "",
        ]
        if stats_lines:
            lines += ["**Column summaries:**"] + stats_lines + [""]
        if extra_notes:
            lines += [f"> {extra_notes}", ""]
        lines += ["**Preview (first 20 rows):**"]
        result["response_text"] = "\n".join(lines)

    return result


def format_mcp_response(smart_result: dict, task: str, code: str = "") -> str:
    """
    Build the final string returned by the MCP tool.
    Embeds base64 images as markdown image tags so LLMs that support it
    can render them inline. Falls back to file paths for others.
    """
    parts = [f"## Task: {task}\n"]

    # Main content
    parts.append(smart_result["response_text"])

    # Chart (always first if present)
    if smart_result["chart_b64"]:
        parts.append(
            f"\n**Chart:**\n"
            f"![chart](data:image/png;base64,{smart_result['chart_b64']})"
        )

    # Thumbnail (for tier 2/3)
    if smart_result["thumbnail_b64"] and smart_result["tier"] > 1:
        parts.append(
            f"\n**Data Preview:**\n"
            f"![preview](data:image/png;base64,{smart_result['thumbnail_b64']})"
        )

    # File links
    if smart_result["html_path"]:
        parts.append(f"\n📄 Open in browser: `{smart_result['html_path']}`")
    if smart_result["excel_path"]:
        parts.append(f"\n📥 Download Excel: `{smart_result['excel_path']}`")

    # Code
    if code:
        parts.append(f"\n### Code executed\n```python\n{code.strip()}\n```")

    return "\n".join(parts)