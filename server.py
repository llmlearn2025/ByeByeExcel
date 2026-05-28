"""
Excel MCP Server v2
===================
One server. One storage layer. LLM holds the intelligence.

Design principles (see architecture diagram in session history):
  - MCP provides: graph (structure), fuzzy primitives, execution engine,
                  storage layer, rich output
  - LLM decides:  which columns to dedup on, which normaliser fits,
                  what thresholds make sense, how to interpret results
  - MCP never:    hardcodes column semantics, decides normaliser selection,
                  assumes beneficiary/budget/any specific domain

V1 files are preserved unchanged. V2 is a clean redesign.

Tools:
  excel_init(path)                        → copy + graph + format scan
  excel_inspect(ai_path)                  → refresh graph
  excel_execute(ai_path, task)            → plain-English analysis
  excel_query(ai_path, query)             → smart tiered output
  excel_run_code(ai_path, code)           → custom Python
  excel_audit(ai_path)                    → data quality (on demand only)
  excel_apply_fixes(ai_path, fixes_json)  → write LLM-proposed fixes
  excel_ingest_large(path, db_path)          → Excel → SQLite for 11L+ files
  excel_bogus_detect(ai_path)               → anomaly detection (on demand only)
"""

import logging
import sys as _sys
import os
from contextlib import asynccontextmanager

_sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

import io, re, json, shutil, textwrap, traceback, base64
from pathlib import Path

# Import FastMCP BEFORE defining lifespan to avoid circular reference
from mcp.server.fastmcp import FastMCP

import openpyxl
import pandas as pd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from config      import PORT, CHART_DPI
from graph       import build_graph, render_graph_md
from storage     import StorageBackend, _chunked_excel_reader
from normalisers import REGISTRY, suggest_for_column, describe_all, get_normaliser
from rich_output    import smart_return, format_mcp_response

from pydantic import BaseModel, Field
class TransportSecuritySettings(BaseModel):
    """Settings for MCP transport security features."""
    enable_dns_rebinding_protection: bool = Field(
        default=True,
        description="Enable DNS rebinding protection (recommended for production)",
    )
    allowed_hosts: list[str] = Field(
        default=["127.0.0.1:*", "localhost:*", "host.docker.internal:*"],
        description="List of allowed Host header values.",
    )
    allowed_origins: list[str] = Field(
        default=["http://127.0.0.1:*", "http://localhost:*"],
        description="List of allowed Origin header values.",
    )

security_config = TransportSecuritySettings(
    enable_dns_rebinding_protection=True,
    allowed_hosts=[
        "127.0.0.1:*", 
        "localhost:*", 
        "host.docker.internal:*", # Allows Docker networking
        "0.0.0.0:*" 
    ],
    allowed_origins=[
        "http://127.0.0.1:*", 
        "http://localhost:*",
        "http://host.docker.internal:*"
    ]
)

# Server lifespan for initialization and cleanup
@asynccontextmanager
async def lifespan(app: FastMCP):
    """Initialize resources on server startup and clean up on shutdown."""
    logger = logging.getLogger(__name__)
    logger.info("Excel MCP Server v2 starting...")
    yield
    logger.info("Excel MCP Server v2 shutting down...")

# Initialize FastMCP with lifespan
mcp = FastMCP(
    name="excel-analyst-v2",
    port=PORT,
    transport_security=security_config,
    lifespan=lifespan,
)

# Set port via environment variable
os.environ["FASTMCP_PORT"] = str(PORT)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _fig_b64() -> str:
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=CHART_DPI, bbox_inches="tight")
    buf.seek(0)
    b64 = base64.b64encode(buf.read()).decode()
    plt.close("all")
    return b64


def _load_graph(ai_path: str) -> dict:
    p = Path(ai_path).resolve()
    graph_path = p.parent / (p.stem.replace("_ai_workbook","") + "_graph.json")
    if graph_path.exists():
        return json.loads(graph_path.read_text())
    return {}


def _save_graph(g: dict, ai_path: str):
    p = Path(ai_path).resolve()
    graph_path = p.parent / (p.stem.replace("_ai_workbook","") + "_graph.json")
    g["graph_path"] = str(graph_path)
    graph_path.write_text(json.dumps(g, indent=2, default=str))
    return str(graph_path)


def _analysis_dir(ai_path: str) -> Path:
    p = Path(ai_path).resolve()
    d = p.parent / (p.stem.replace("_ai_workbook","") + "_analysis")
    d.mkdir(exist_ok=True)
    return d


def _make_exec_env(ai_path: str, graph: dict, out_dir: Path) -> dict:
    """Execution sandbox for excel_run_code and excel_execute."""
    try: import numpy as np
    except ImportError: np = None
    outputs = []

    def save_chart(name="chart"):
        plt.tight_layout()
        path = out_dir / f"{name}.png"
        plt.savefig(str(path), dpi=CHART_DPI, bbox_inches="tight")
        plt.close("all")
        return str(path)

    def save_new_workbook(df, sheet_name="Analysis", fname=None):
        fpath = out_dir / (fname or f"{sheet_name}.xlsx")
        if isinstance(df.columns, pd.MultiIndex):
            df = df.copy()
            df.columns = ["_".join(str(c) for c in col).strip("_")
                          for col in df.columns]
        with pd.ExcelWriter(str(fpath), engine="openpyxl") as w:
            df.to_excel(w, sheet_name=sheet_name[:31], index=True)
        return str(fpath)

    def append_sheet(df, sheet_name="AI_Analysis"):
        wb = openpyxl.load_workbook(ai_path)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        wb.save(ai_path)
        with pd.ExcelWriter(ai_path, engine="openpyxl", mode="a") as w:
            df.to_excel(w, sheet_name=sheet_name[:31], index=False)
        return f"Sheet '{sheet_name}' added to {ai_path}"

    def read_sheet(sheet=0, nrows=None):
        return pd.read_excel(ai_path, sheet_name=sheet, nrows=nrows)

    def read_sheet_chunked(sheet=0, chunksize=50000):
        return _chunked_excel_reader(ai_path, sheet, chunksize)

    env = {
        "pd": pd, "plt": plt, "openpyxl": openpyxl,
        "Path": Path, "json": json, "re": re, "np": np,
        "FILE_PATH":    ai_path,
        "AI_WORKBOOK":  ai_path,
        "OUTPUT_DIR":   str(out_dir),
        "GRAPH":        graph,
        "read_sheet":            read_sheet,
        "read_sheet_chunked":    read_sheet_chunked,
        "read_all_sheets":       lambda: pd.read_excel(ai_path, sheet_name=None),
        "save_chart":            save_chart,
        "save_new_workbook":     save_new_workbook,
        "append_sheet":          append_sheet,
        "_outputs":              outputs,
        "print": lambda *a, **kw: outputs.append(" ".join(str(x) for x in a)),
    }
    return env


# ─────────────────────────────────────────────────────────────────────────────
# TOOLS
# ─────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_init(path: str) -> str:
    """
    Initialise the AI workspace for an Excel file.
    Call this once per file before any other tool.

    Creates:
      <stem>_ai_workbook.xlsx  — working copy (all reads/writes go here)
      <stem>_graph.json        — full structural + formatting graph
      <stem>_analysis/         — output directory for charts, exports

    The original file is NEVER touched after this point.

    Returns the complete graph as markdown — the LLM must read this
    before calling any other tool.
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"
    if p.suffix.lower() not in (".xlsx",".xlsm",".xlam",".xls"):
        return f"ERROR: Not an Excel file: {path}"

    ai_path  = p.parent / f"{p.stem}_ai_workbook{p.suffix}"
    out_dir  = p.parent / f"{p.stem}_analysis"
    out_dir.mkdir(exist_ok=True)

    try:
        shutil.copy2(str(p), str(ai_path))
        g = build_graph(str(ai_path), original_path=str(p))
        g["output_dir"] = str(out_dir)
        graph_path = _save_graph(g, str(ai_path))

        md = render_graph_md(g)

        # Check for existing skill document
        from workbook_skill import get_skill_summary, init_skill
        skill_summary = get_skill_summary(str(ai_path))
        if not skill_summary["exists"]:
            init_skill(str(ai_path))
            skill_note = "*(new — call `excel_update_skill` to add domain context)*"
        else:
            parts = []
            if skill_summary.get("context_preview"):
                parts.append(skill_summary["context_preview"][:80] + "…")
            if skill_summary.get("rule_count",0) > 0:
                parts.append(f"{skill_summary['rule_count']} rules")
            if skill_summary.get("warning_count",0) > 0:
                parts.append(f"{skill_summary['warning_count']} warnings")
            if skill_summary.get("linked_queries",0) > 0:
                parts.append(f"{skill_summary['linked_queries']} linked queries")
            skill_note = " | ".join(parts) if parts else "*(exists — call `excel_read_skill` to load)*"

        from knowledge import get_memory_stats
        mem_stats = get_memory_stats(str(ai_path))
        mem_note = (f"{mem_stats['count']} saved queries"
                    if mem_stats.get("count",0) > 0 else "*(none yet)*")

        return (
            md +
            f"\n\n---\n## Workspace Ready\n"
            f"- **Work on:** `{ai_path}`\n"
            f"- **Graph:** `{graph_path}`\n"
            f"- **Outputs:** `{out_dir}`\n"
            f"- **Original (preserved):** `{p}`\n\n"
            f"## Accumulated Knowledge\n"
            f"- **Skill document:** {skill_note}\n"
            f"- **Query memory:** {mem_note}\n\n"
            f"*Call `excel_read_skill` to load domain context.*\n"
            f"*Call `excel_search_docs` to find relevant patterns.*\n"
        )
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_inspect(path: str) -> str:
    """
    Refresh the workbook graph. Use after manual edits or after
    excel_apply_fixes changes the workbook.

    Args:
        path: Path to _ai_workbook.xlsx
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"
    try:
        g = build_graph(str(p))
        _save_graph(g, str(p))
        return render_graph_md(g) + f"\n\n*Graph refreshed.*"
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_get_fuzzy_options(path: str, columns: str) -> str:
    """
    Ask the MCP what fuzzy matching strategies are available and which
    fit the specified columns. The LLM reads this, then decides which
    normaliser to pass to excel_fuzzy_dedup.

    The MCP never decides — it informs. The LLM chooses.

    Args:
        path:    Path to _ai_workbook.xlsx (graph must exist)
        columns: Comma-separated column names to get suggestions for
                 e.g. "Name,Father_Name,Village"

    Returns:
      - Description of all available normalisers with examples
      - Per-column suggestions with reasoning based on graph data
        (column format_meaning, fill_meaning, sample values)
    """
    p = Path(path).resolve()
    g = _load_graph(str(p))

    col_list = [c.strip() for c in columns.split(",") if c.strip()]
    if not col_list:
        # Return full registry description
        all_desc = describe_all()
        lines = ["## Available Fuzzy Normalisers\n",
                 "Pass normaliser names to `excel_fuzzy_dedup` via `normalisers` parameter.\n"]
        for name, info in all_desc.items():
            lines.append(f"### `{name}`")
            lines.append(info["description"])
            lines.append(f"- **Best for:** {', '.join(info['best_for'][:2])}")
            lines.append(f"- **Not for:** {', '.join(info['not_for'][:1])}")
            if info.get("example"):
                ex = info["example"]
                if "input_a" in ex:
                    lines.append(
                        f"- **Example:** `{ex.get('input_a')}` vs "
                        f"`{ex.get('input_b')}` → score={ex.get('score','?')}")
            lines.append("")
        return "\n".join(lines)

    # Per-column suggestions using graph data
    lines = ["## Normaliser Suggestions by Column\n",
             "Review and pass your chosen normalisers to `excel_fuzzy_dedup`.\n"]

    for col in col_list:
        # Get column data from graph
        col_type = "unknown"
        fmt_meaning = "general"
        fill_meaning = "none"
        sample_vals = []

        for sname, s_info in g.get("sheets", {}).items():
            if col in (s_info.get("column_types") or {}):
                col_type = s_info["column_types"][col]
                fmt_info = s_info.get("column_formats", {}).get(col, {})
                fmt_meaning  = fmt_info.get("format_meaning", "general")
                fill_meaning = fmt_info.get("fill_meaning", "none")
                # Pull sample values
                for row in s_info.get("sample_rows", []):
                    if col in row and row[col] is not None:
                        sample_vals.append(row[col])
                break

        suggestions = suggest_for_column(col, sample_vals, fmt_meaning, fill_meaning)

        lines.append(f"### Column: `{col}`")
        lines.append(f"- Graph says: type=`{col_type}`, "
                     f"format=`{fmt_meaning}`, fill=`{fill_meaning}`")
        if sample_vals:
            lines.append(f"- Sample values: {sample_vals[:4]}")
        lines.append("- **Suggestions:**")
        for s in suggestions:
            conf_icon = {"high":"✅","medium":"⚠️","low":"ℹ️"}.get(s["confidence"],"")
            lines.append(f"  {conf_icon} `{s['normaliser']}` "
                         f"({s['confidence']} confidence): {s['reason']}")
        lines.append("")

    cols_str = ",".join(col_list)
    norms_str = ",".join(
        suggest_for_column(c, [], "", "")[0]["normaliser"] for c in col_list)
    lines += [
        "## How to use these suggestions",
        "```",
        "excel_fuzzy_dedup(",
        "    path='...ai_workbook.xlsx',",
        f"    columns='{cols_str}',",
        f"    normalisers='{norms_str}',  # adjust based on suggestions above",
        "    weights='0.6,0.25,0.15',",
        "    high_threshold=0.80,",
        "    review_threshold=0.60,",
        "    possible_threshold=0.40,",
        ")",
        "```",
    ]
    return "\n".join(lines)


@mcp.tool()
def excel_execute(path: str, task: str) -> str:
    """
    Execute a plain-English analysis task on the ai_workbook.
    Runs server-side. Returns results + charts. Safe for large files.

    Task examples:
      "summarise all sheets"
      "chart Revenue by Month"
      "pivot Units_Sold by Region"
      "total all numeric columns"
      "count duplicates"
      "show missing values"
      "add summary as new sheet"

    Args:
        path: Path to _ai_workbook.xlsx
        task: Plain English description
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"
    out = _analysis_dir(str(p))

    try:
        g = _load_graph(str(p))
        sheets = list(g.get("sheets", {}).keys())
        fs = sheets[0] if sheets else "Sheet1"

        s_info   = g.get("sheets", {}).get(fs, {})
        headers  = [str(h) for h in s_info.get("headers",[]) if h is not None]
        ctypes   = s_info.get("column_types", {})
        cfmts    = s_info.get("column_formats", {})
        numeric  = [k for k,v in ctypes.items() if v=="numeric"
                    and not cfmts.get(k,{}).get("hidden")]
        text_c   = [k for k,v in ctypes.items() if v=="text"]
        pct_c    = [k for k,v in cfmts.items()
                    if v.get("format_meaning")=="percentage"]
        max_row  = s_info.get("max_row", 0)
        big      = max_row > 10000
        t        = task.lower()

        # Generate code for the task
        if any(w in t for w in ["summary","overview","describe","stat"]):
            code = textwrap.dedent(f"""
                df = read_sheet("{fs}")
                print("Shape:", df.shape)
                print("\\nColumn types:"); print(df.dtypes.to_string())
                print("\\nNumeric summary:"); print(df.describe().to_string())
                print("\\nNull counts:"); print(df.isnull().sum().to_string())
                if {pct_c}:
                    print("\\nNote: percentage columns {pct_c} are 0–1 decimals")
            """)
        elif any(w in t for w in ["chart","graph","plot","visual"]):
            xc = text_c[0] if text_c else (headers[0] if headers else None)
            yc = numeric[0] if numeric else None
            if xc and yc:
                pct_fmt = (f'ax.yaxis.set_major_formatter('
                           f'plt.FuncFormatter(lambda x,_: f"{{x*100:.1f}}%"))'
                           if yc in pct_c else "")
                code = textwrap.dedent(f"""
                    df = read_sheet("{fs}")
                    grp = df.groupby("{xc}")["{yc}"].sum().reset_index()
                    grp = grp.sort_values("{yc}", ascending=False)
                    fig, ax = plt.subplots(figsize=(12, 6))
                    ax.barh(grp["{xc}"].astype(str), grp["{yc}"], color="#4472C4")
                    ax.set_xlabel("{yc}"); ax.set_title("{yc} by {xc}", fontsize=12)
                    {pct_fmt}
                    plt.tight_layout()
                    print("Chart saved:", save_chart("chart_{yc}_by_{xc}"))
                """)
            else:
                code = 'print("Need at least one text and one numeric column.")'
        elif any(w in t for w in ["pivot","group","breakdown","aggregate"]):
            if numeric and text_c:
                code = textwrap.dedent(f"""
                    df = read_sheet("{fs}")
                    pivot = df.groupby("{text_c[0]}")[{str(numeric[:3])}].agg(["sum","mean","count"])
                    print(pivot.to_string())
                    path = save_new_workbook(pivot.reset_index(), "Pivot_{text_c[0]}")
                    print("Saved:", path)
                """)
            else:
                code = 'print("Need text and numeric columns for pivot.")'
        elif any(w in t for w in ["total","sum"]):
            cols = str(numeric[:5])
            if big:
                code = textwrap.dedent(f"""
                    totals = {{c: 0 for c in {cols}}}; n = 0
                    for chunk in read_sheet_chunked("{fs}", 50000):
                        for c in {cols}:
                            if c in chunk.columns:
                                totals[c] += chunk[c].fillna(0).sum()
                        n += len(chunk)
                    print(f"Processed {{n:,}} rows")
                    for k, v in totals.items():
                        sfx = "%" if k in {pct_c} else ""
                        display = f"{{v*100:.2f}}%" if k in {pct_c} else f"{{v:,.2f}}"
                        print(f"  {{k}}: {{display}}")
                """)
            else:
                code = textwrap.dedent(f"""
                    df = read_sheet("{fs}")
                    for c in {cols}:
                        if c in df.columns:
                            v = df[c].sum()
                            display = f"{{v*100:.2f}}%" if c in {pct_c} else f"{{v:,.2f}}"
                            print(f"  {{c}}: {{display}}")
                """)
        elif any(w in t for w in ["missing","null","blank"]):
            code = textwrap.dedent(f"""
                df = read_sheet("{fs}")
                nulls = df.isnull().sum(); pct = (nulls/len(df)*100).round(2)
                rep = __import__("pandas").DataFrame({{"nulls":nulls,"pct%":pct}})
                rep = rep[rep["nulls"]>0]
                print(rep.to_string() if len(rep) else "No missing values.")
            """)
        elif any(w in t for w in ["duplicate","dedup"]):
            code = textwrap.dedent(f"""
                df = read_sheet("{fs}")
                dups = df.duplicated().sum()
                print(f"Rows: {{len(df):,}} | Exact duplicates: {{dups:,}}")
                if dups: print(df[df.duplicated(keep=False)].head(5).to_string())
                print("\\nNote: use excel_fuzzy_dedup for fuzzy/near-duplicate detection")
            """)
        elif any(w in t for w in ["all sheet","every sheet","each sheet"]):
            code = textwrap.dedent(f"""
                sheets = read_all_sheets()
                for name, df in sheets.items():
                    print(f"\\n=== {{name}} === ({{len(df):,}} rows x {{len(df.columns)}} cols)")
                    print("  Headers:", list(df.columns))
                    num = df.select_dtypes(include="number")
                    if not num.empty:
                        print("  Totals:", {{c: f"{{df[c].sum():,.0f}}" for c in num.columns[:4]}})
            """)
        elif any(w in t for w in ["new sheet","add sheet","append"]):
            if numeric and text_c:
                code = textwrap.dedent(f"""
                    df = read_sheet("{fs}")
                    summary = df.groupby("{text_c[0]}")[{str(numeric[:3])}].sum().reset_index()
                    result = append_sheet(summary, "AI_Summary")
                    print(result)
                """)
            else:
                code = 'print("Need text and numeric columns for summary sheet.")'
        else:
            code = textwrap.dedent(f"""
                df = read_sheet("{fs}")
                print(f"Sheet: {fs} | {{len(df):,}} rows × {{len(df.columns)}} cols")
                print("Headers:", list(df.columns))
                num = df.select_dtypes(include="number")
                if not num.empty: print("\\nSummary:"); print(num.describe().to_string())
            """)

        env = _make_exec_env(str(p), g, out)
        exec(compile(code, "<task>", "exec"), env)

        text_out = "\n".join(env["_outputs"])
        files = sorted(out.glob("*.png")) + sorted(out.glob("*.xlsx"))
        file_list = "\n".join(f"  📁 {f}" for f in files) or "  (none)"

        # Embed any chart created
        chart_b64 = None
        pngs = sorted(out.glob("*.png"))
        if pngs:
            img = pngs[-1].read_bytes()
            chart_b64 = base64.b64encode(img).decode()

        resp = (f"## Task: {task}\n\n"
                f"### Output\n```\n{text_out or '(no output)'}\n```\n\n"
                f"### Files\n{file_list}\n\n"
                f"### Code\n```python\n{code.strip()}\n```")
        if chart_b64:
            resp += f"\n\n![chart](data:image/png;base64,{chart_b64})"
        return resp

    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_run_code(path: str, code: str) -> str:
    """
    Run custom Python against the ai_workbook. Full control.

    Pre-loaded:
        pd, np, plt, openpyxl, Path, json, re
        FILE_PATH / AI_WORKBOOK   → ai_workbook path
        OUTPUT_DIR                → analysis directory
        GRAPH                     → full graph dict
        read_sheet(sheet, nrows)
        read_sheet_chunked(sheet, chunksize=50000)
        read_all_sheets()
        save_chart(name)          → saves plt figure, returns path
        save_new_workbook(df, sheet_name, fname)
        append_sheet(df, sheet_name)
        print(...)                → captured and returned

    For files >50K rows, always use read_sheet_chunked:
        for chunk in read_sheet_chunked("Sheet1", 50000):
            process(chunk)

    Args:
        path: Path to _ai_workbook.xlsx
        code: Python code to run
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"
    out = _analysis_dir(str(p))
    try:
        g   = _load_graph(str(p))
        env = _make_exec_env(str(p), g, out)
        exec(compile(code, "<custom>", "exec"), env)
        text_out = "\n".join(env["_outputs"])
        files = sorted(out.glob("*.png")) + sorted(out.glob("*.xlsx"))
        file_list = "\n".join(f"  📁 {f}" for f in files) or "  (none)"

        chart_b64 = None
        pngs = sorted(out.glob("*.png"))
        if pngs:
            chart_b64 = base64.b64encode(pngs[-1].read_bytes()).decode()

        resp = (f"### Output\n```\n{text_out or '(no output)'}\n```\n\n"
                f"### Files\n{file_list}")
        if chart_b64:
            resp += f"\n\n![chart](data:image/png;base64,{chart_b64})"
        return resp
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_audit(path: str) -> str:
    """
    Data quality audit. Call ONLY when user explicitly asks for it.
    Never runs automatically at init.

    Detects (same as v1 audit.py):
      CRITICAL  FORMULA_OVERRIDE  — hardcoded value where formula expected
      HIGH      LOGIC_VIOLATION   — SUM(parts) ≠ total
      HIGH      TYPE_VIOLATION    — text in numeric column
      MEDIUM    FORMAT_VIOLATION  — 75 in a 0.00% column instead of 0.75
      MEDIUM    RANGE_VIOLATION   — negative budget, utilisation > 100%
      MEDIUM    MISSING_REQUIRED  — null in otherwise-full column
      LOW       DUPLICATE_ROW     — exact duplicate rows

    Additionally activates bogus-pattern checks when canonical columns
    detected in the graph (aadhaar, mobile, bank_account):
      B01  Aadhaar checksum failure
      B02  Aadhaar fake pattern
      B03  Mobile fake pattern
      B04  Mobile shared by 3+ names
      B05  Aadhaar shared by 2+ names
      B06  Bank account shared by 3+ names
      B07  Name = Father name

    Args:
        path: Path to _ai_workbook.xlsx
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"
    out = _analysis_dir(str(p))

    try:
        from audit import audit_workbook, render_audit_html, render_audit_md
        g = _load_graph(str(p))
        if not g:
            g = build_graph(str(p))

        audit_result = audit_workbook(str(p), g)

        # Check if bogus detection is relevant
        # (graph has aadhaar or mobile columns)
        has_id_cols = any(
            any(f"_c_{field}" in str(s.get("column_formats", {}))
                or field in str(s.get("column_types", {}))
                for field in ["aadhaar","mobile","bank_account"])
            for s in g.get("sheets", {}).values()
        )

        bogus_note = ""
        if has_id_cols:
            bogus_note = (
                "\n\n> **Note:** Identity columns detected "
                "(Aadhaar/mobile/bank account). "
                "Run `excel_bogus_detect` for additional "
                "fraud/anomaly checks on those columns."
            )

        stem      = p.stem.replace("_ai_workbook","")
        json_out  = out / f"{stem}_audit_issues.json"
        html_out  = out / f"{stem}_audit_report.html"

        json_out.write_text(
            json.dumps(audit_result, indent=2, default=str))
        render_audit_html(audit_result, stem, html_out)

        # Chart
        chart_b64 = None
        by_type   = audit_result["summary"].get("by_type", {})
        if by_type:
            plt.close("all")
            fig, axes = plt.subplots(1, 2, figsize=(11, 4))
            sev_data = {s: audit_result["summary"]["by_severity"].get(s,0)
                        for s in ["CRITICAL","HIGH","MEDIUM","LOW"]}
            sev_data = {k:v for k,v in sev_data.items() if v}
            sc = {"CRITICAL":"#c0392b","HIGH":"#e67e22",
                  "MEDIUM":"#f39c12","LOW":"#27ae60"}
            if sev_data:
                axes[0].barh(list(sev_data.keys()), list(sev_data.values()),
                            color=[sc[k] for k in sev_data])
                axes[0].set_title("By severity", fontweight="bold")
            ts = sorted(by_type.items(), key=lambda x: -x[1])[:8]
            if ts:
                axes[1].barh([t for t,_ in ts],[c for _,c in ts],color="#4472C4")
                axes[1].set_title("By type", fontweight="bold")
            plt.suptitle(f"Audit: {stem}", fontsize=11, fontweight="bold")
            plt.tight_layout()
            chart_b64 = _fig_b64()

        md = render_audit_md(audit_result, str(html_out), str(json_out))
        md += bogus_note
        if chart_b64:
            md += f"\n\n![audit chart](data:image/png;base64,{chart_b64})"
        return md

    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_apply_fixes(path: str, fixes_json: str) -> str:
    """
    Apply LLM-proposed fixes to the ai_workbook.
    Writes to ai_workbook only — original never touched.

    fixes_json: JSON string or path to .json file containing a list:
    [
      {"sheet": "Sheet1", "row": 5, "col_letter": "K",
       "new_formula": "=J5/E5"},
      {"sheet": "Sheet1", "row": 12, "col_letter": "E",
       "new_value": 5000000}
    ]

    Args:
        path:       Path to _ai_workbook.xlsx
        fixes_json: JSON fixes list (string) or path to a .json file
    """
    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"

    try:
        if fixes_json.strip().startswith("[") or fixes_json.strip().startswith("{"):
            fixes = json.loads(fixes_json)
        else:
            fpath = Path(fixes_json)
            if not fpath.exists(): return f"ERROR: File not found: {fixes_json}"
            fixes = json.loads(fpath.read_text())

        if not isinstance(fixes, list):
            fixes = [fixes]

        from audit import apply_fixes
        result = apply_fixes(str(p), fixes)

        # Invalidate graph cache
        graph_path = p.parent / (p.stem.replace("_ai_workbook","") + "_graph.json")
        if graph_path.exists():
            graph_path.unlink()

        return (f"## Fixes Applied\n\n"
                f"```\n{result['diff_summary']}\n```\n\n"
                f"Graph cache cleared. Run `excel_inspect` to refresh.")

    except json.JSONDecodeError as e:
        return f"ERROR: Invalid JSON: {e}"
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_ingest_large(
    excel_path:   str,
    db_path:      str      = "",
    overwrite:    bool     = False,
    merge_sheets: bool     = False,
) -> str:
    """
    Ingest an Excel file into SQLite for large-file workflows (>50K rows).

    Default behaviour (merge_sheets=False):
      Creates one SQLite table per sheet, named from the sanitised sheet name.
        "Expenditure"  → table  expenditure
        "Revenue Data" → table  revenue_data
      Use this for workbooks where sheets have different schemas
      (budget files, multi-entity workbooks).

    merge_sheets=True:
      All sheets combined into one table named 'data'.
      Use this when the same schema is split across sheets
      (e.g. an 11-lakh beneficiary register split as Sheet1/Sheet2/Sheet3).

    Freshness check:
      If the db already exists and the Excel file has not been modified
      since the last ingest, returns "DB is current" and skips re-ingestion.
      Pass overwrite=True to force a full rebuild regardless.

    Args:
        excel_path:   Path to the Excel file (original or ai_workbook)
        db_path:      Output .db path. Defaults to same directory as Excel,
                      same stem, .db extension.
        overwrite:    Drop and recreate even if db is current (default False)
        merge_sheets: Combine all sheets into one table (default False)

    Returns:
      Ingestion report including sheet_tables mapping, SQL JOIN example,
      and freshness check result. The LLM can use sheet_tables directly
      to write SQL queries against the correct table for each sheet.
    """
    p  = Path(excel_path)
    if not p.exists():
        return f"ERROR: File not found: {excel_path}"
    db = db_path or str(p.with_suffix(".db"))

    try:
        from ingest import ingest_excel_to_sqlite
        result = ingest_excel_to_sqlite(
            excel_path,
            db,
            overwrite    = overwrite,
            merge_sheets = merge_sheets,
        )

        # ── Already current ───────────────────────────────────────────────────
        if result.get("status") == "current":
            sheet_tables = result.get("sheet_tables", {})
            table_lines  = _format_sheet_tables(sheet_tables)
            return (
                f"## DB is Current\n\n"
                f"{result['message']}\n\n"
                f"- **Database:** `{db}`\n"
                f"- **Total rows:** {result['total_rows']:,}\n\n"
                f"### Sheet → Table Mapping\n{table_lines}\n\n"
                f"{_sql_join_example(sheet_tables)}"
            )

        # ── Fresh ingest ──────────────────────────────────────────────────────
        sheet_tables = result.get("sheet_tables", {})
        table_lines  = _format_sheet_tables(sheet_tables)

        return (
            f"## Ingestion Complete\n\n"
            f"- **Database:** `{db}`\n"
            f"- **Total rows:** {result['total_rows']:,}\n"
            f"- **Mode:** {'merge_sheets (one table)' if merge_sheets else 'per-sheet tables'}\n\n"
            f"### Sheet → Table Mapping\n{table_lines}\n\n"
            f"{_sql_join_example(sheet_tables)}\n\n"
            f"The database is ready for `excel_fuzzy_dedup`, "
            f"`excel_bogus_detect`, and direct SQL via `excel_run_code`."
        )

    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


def _format_sheet_tables(sheet_tables: dict) -> str:
    """Format sheet_tables dict as markdown for LLM consumption."""
    if not sheet_tables:
        return "*No sheet tables available.*"
    lines = []
    for sheet, info in sheet_tables.items():
        cols     = info.get("columns", [])
        jc       = info.get("join_candidates", [])
        canon    = info.get("canonical_cols", {})
        col_str  = ", ".join(
            f"`{c['name']}` ({c['type']})" for c in cols[:6]
        )
        if len(cols) > 6:
            col_str += f" ... +{len(cols)-6} more"
        jc_str   = (f"\n  - Join keys: {', '.join(f'`{c}`' for c in jc)}"
                    if jc else "")
        can_str  = (f"\n  - Canonical: "
                    + ", ".join(f"`_c_{k}` = `{v}`" for k, v in canon.items())
                    if canon else "")
        lines.append(
            f"- **`{sheet}`** → table `{info['table']}` "
            f"({info['rows']:,} rows)\n"
            f"  - Columns: {col_str}"
            f"{jc_str}{can_str}"
        )
    return "\n".join(lines)


def _sql_join_example(sheet_tables: dict) -> str:
    """Generate a SQL JOIN example from sheet_tables join_candidates."""
    sheets = list(sheet_tables.keys())
    if len(sheets) < 2:
        return ""
    s1, s2 = sheets[0], sheets[1]
    t1     = sheet_tables[s1]["table"]
    t2     = sheet_tables[s2]["table"]
    jc1    = sheet_tables[s1].get("join_candidates", [])
    jk     = jc1[0] if jc1 else "shared_column"

    # Example columns from each table
    c1_cols = [c["name"] for c in sheet_tables[s1].get("columns", [])
               if not c["name"].startswith("_")][:3]
    c2_cols = [c["name"] for c in sheet_tables[s2].get("columns", [])
               if not c["name"].startswith("_")][:2]
    sel     = (", ".join(f"a.[{c}]" for c in c1_cols) +
               (", " + ", ".join(f"b.[{c}]" for c in c2_cols) if c2_cols else ""))

    return (
        f"### SQL JOIN Example\n"
        f"Use `GRAPH['sheet_tables']` in `excel_run_code` for table names:\n"
        f"```python\n"
        f"import sqlite3, pandas as pd\n"
        f"conn = sqlite3.connect(GRAPH['db_path'])\n"
        f"df = pd.read_sql(\"\"\"\n"
        f"    SELECT {sel}\n"
        f"    FROM [{t1}] a\n"
        f"    LEFT JOIN [{t2}] b ON a.[{jk}] = b.[{jk}]\n"
        f"\"\"\", conn)\n"
        f"conn.close()\n"
        f"print(df.head().to_string())\n"
        f"```"
    )



# ─────────────────────────────────────────────────────────────────────────────
# KNOWLEDGE + QUERY MEMORY TOOLS
# ─────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_search_docs(
    query: str,
    path: str = "",
    category: str = "",
    difficulty: str = "",
    top_k: int = 8,
) -> str:
    """
    Search the knowledge base for documentation and working code examples.

    Call this when you don't know how to do something, or want to check
    if a better approach exists. Returns:
      1. Previously working queries for THIS workbook (from query memory)
      2. Relevant documentation entries ranked by BM25 similarity

    The LLM should use returned code as templates — adapt column names
    and sheet names from the graph.

    After successfully using code from search results, call excel_save_query
    to persist it to the workbook's memory for future sessions.

    Args:
        query:      Natural language description of what you need.
                    Examples:
                      "rolling 12-month average"
                      "find rows above average in a column"
                      "budget vs actual comparison chart"
                      "group by district and sum amount"
                      "handle percentage columns correctly"
                      "waterfall chart"
                      "naga name deduplication"
        path:       Path to ai_workbook (optional — enables workbook memory search)
        category:   Filter by: pandas|sql|analysis|visualisation|excel|dedup
        difficulty: Filter by: beginner|intermediate|advanced
        top_k:      Number of doc results (default 8)
    """
    from knowledge import (search_docs, search_query_memory,
                               render_search_results, init_knowledge_db,
                               record_access)

    db_path = str(Path(__file__).parent / "excel_knowledge.db")
    if not Path(db_path).exists():
        init_knowledge_db(db_path)
        return (
            "Knowledge base is empty. Run the seed script first:\n"
            "```\npython seed_corpus.py\n```\n"
            "Or add entries via `excel_add_doc_entry`."
        )

    try:
        doc_results = search_docs(
            query,
            category   = category or None,
            difficulty = difficulty or None,
            top_k      = top_k,
            db_path    = db_path,
        )

        # Record access for top results (boosts them in future searches)
        for r in doc_results[:3]:
            record_access(r["id"], db_path=db_path)

        memory_results = []
        if path:
            p = Path(path).resolve()
            if p.exists():
                memory_results = search_query_memory(str(p), query, top_k=5)

        return render_search_results(doc_results, memory_results, query, path)

    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_save_query(
    path: str,
    description: str,
    code: str,
    sheet: str = "",
    columns_used: str = "",
    result_summary: str = "",
    tags: str = "",
) -> str:
    """
    Save a successful query to the workbook's memory.

    Call this after ANY successful analysis so future sessions can reuse it.
    The MCP deduplicates — running the same query twice just increments
    the execution count. High-count queries appear first in future searches.

    This is the self-improving loop:
      1. LLM searches docs → finds template → adapts it
      2. Analysis succeeds
      3. LLM calls excel_save_query with the working code
      4. Next session: workbook memory surfaces this first, before generic docs

    Args:
        path:           Path to ai_workbook.xlsx
        description:    Plain English description (used for future search matching)
                        e.g. "utilisation by department as horizontal bar chart"
        code:           The working Python code
        sheet:          Sheet name the code targets
        columns_used:   Comma-separated column names used in the code
        result_summary: Brief description of what the result showed
                        e.g. "Agriculture 87%, Health 45%, Education 62%"
        tags:           Comma-separated tags for better search matching
                        e.g. "utilisation,bar chart,department"
    """
    from knowledge import save_query

    p = Path(path).resolve()
    if not p.exists(): return f"ERROR: File not found: {path}"

    try:
        cols = [c.strip() for c in columns_used.split(",") if c.strip()]
        tag_list = [t.strip() for t in tags.split(",") if t.strip()]

        entry = save_query(
            ai_path        = str(p),
            description    = description,
            code           = code,
            sheet          = sheet,
            columns_used   = cols,
            result_summary = result_summary,
            tags           = tag_list,
        )

        return (
            f"✅ Query saved to workbook memory.\n\n"
            f"- **Description:** {entry['description']}\n"
            f"- **Key:** `{entry['key']}`\n"
            f"- **Execution count:** {entry['execution_count']}\n"
            f"- **Sheet:** `{entry.get('sheet','')}`\n"
            f"- **Columns:** {', '.join(f'`{c}`' for c in entry.get('columns_used',[]))}\n\n"
            f"This will appear at the top of future `excel_search_docs` results "
            f"for this workbook."
        )
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_add_doc_entry(
    title: str,
    problem: str,
    code: str,
    category: str = "pandas",
    subcategory: str = "",
    tags: str = "",
    notes: str = "",
    source: str = "",
    difficulty: str = "intermediate",
) -> str:
    """
    Add a new entry to the shared knowledge base.

    Use this to:
      - Add a working pattern discovered during analysis
      - Populate the corpus from textbook content
      - Share useful patterns across workbooks (unlike excel_save_query
        which is per-workbook, this goes into the global knowledge base)

    The entry is immediately searchable via excel_search_docs.

    Args:
        title:       Short descriptive title (used in search ranking)
        problem:     Plain English description of what problem this solves
        code:        Working Python/SQL code
        category:    pandas|sql|analysis|visualisation|excel|dedup
        subcategory: groupby|pivot|filter|window|date|chart|formula|...
        tags:        Comma-separated keywords for search
        notes:       Gotchas, limitations, when NOT to use
        source:      Book title, chapter, URL, or "discovered in analysis"
        difficulty:  beginner|intermediate|advanced
    """
    from knowledge import insert_entry, init_knowledge_db

    db_path = str(Path(__file__).parent / "excel_knowledge.db")
    init_knowledge_db(db_path)

    try:
        tag_list = [t.strip() for t in tags.split(",") if t.strip()]
        eid = insert_entry({
            "category":    category,
            "subcategory": subcategory,
            "title":       title,
            "problem":     problem,
            "tags":        tag_list,
            "code":        code,
            "notes":       notes,
            "source":      source,
            "difficulty":  difficulty,
        }, db_path)
        return (
            f"✅ Entry added to knowledge base.\n"
            f"- **ID:** `{eid}`\n"
            f"- **Title:** {title}\n"
            f"- **Category:** {category}/{subcategory}\n"
            f"- Immediately searchable via `excel_search_docs`."
        )
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_kb_stats(detail: bool = False) -> str:
    """
    Show knowledge base statistics: total entries, breakdown by category,
    top accessed entries.

    Args:
        detail: If True, show top 10 accessed entries
    """
    from knowledge import db_stats, init_knowledge_db

    db_path = str(Path(__file__).parent / "excel_knowledge.db")
    if not Path(db_path).exists():
        init_knowledge_db(db_path)
        return "Knowledge base is empty. Run: `python seed_corpus.py`"

    try:
        s = db_stats(db_path)
        lines = [
            "## Knowledge Base Statistics\n",
            f"- **Total entries:** {s['total_entries']:,}",
            f"- **Path:** `{s['path']}`\n",
            "### By Category",
        ]
        for cat, cnt in sorted(s.get("by_category",{}).items(), key=lambda x: -x[1]):
            lines.append(f"  - `{cat}`: {cnt}")
        lines.append("\n### By Difficulty")
        for diff, cnt in s.get("by_difficulty",{}).items():
            lines.append(f"  - `{diff}`: {cnt}")
        if detail and s.get("top_accessed"):
            lines.append("\n### Most Accessed Entries")
            for e in s["top_accessed"]:
                lines.append(f"  - {e['title']} (accessed {e['count']}×)")
        return "\n".join(lines)
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


# ─────────────────────────────────────────────────────────────────────────────
# WORKBOOK SKILL TOOLS
# ─────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_read_skill(path: str) -> str:
    """
    Read the workbook skill document — the LLM's accumulated analytical
    understanding of this specific workbook.

    Returns:
      - Full skill document (Context, Key Columns, Analytical Rules,
        Useful Analyses, Warnings, User Preferences, History)
      - Referenced working code from queries.json inline

    Call this at the start of any session where you already worked with
    this workbook before. The skill document tells you domain context,
    validated rules, and which analyses have worked, saving you from
    rediscovering the same things.

    Args:
        path: Path to _ai_workbook.xlsx
    """
    from workbook_skill import render_skill_for_llm, skill_exists

    p = Path(path).resolve()
    if not p.exists():
        return f"ERROR: File not found: {path}"
    return render_skill_for_llm(str(p), include_queries=True)


@mcp.tool()
def excel_update_skill(
    path: str,
    section: str,
    content: str,
    mode: str = "append",
    query_key: str = "",
) -> str:
    """
    Write to the workbook skill document. Call this whenever you discover
    something important about a workbook that should persist across sessions.

    The skill document is the LLM's analytical memory for this workbook.
    It is distinct from queries.json (which stores code) — the skill document
    stores conceptual understanding, rules, and validated patterns.

    When to call:
      - After learning what the workbook is for → update "Context"
      - After discovering a column quirk → update "Key Columns"
      - After validating an analysis approach → update "Useful Analyses"
      - After hitting a problem (formula breaks, hidden column) → update "Warnings"
      - After user expresses a preference → update "User Preferences"

    Args:
        path:      Path to _ai_workbook.xlsx
        section:   Section to update. One of:
                     Context | Key Columns | Analytical Rules |
                     Useful Analyses | Warnings | User Preferences
        content:   Text to add. Markdown supported.
                   For "Useful Analyses", include query key references:
                   "- **Utilisation Chart**: barh of K by B [query: abc123def]"
        mode:      append (default) | replace | prepend
        query_key: Optional: link this update to a queries.json key

    Examples:
        # After learning the file is a budget tracker:
        excel_update_skill(path, "Context",
            "FY 2025-26 Nagaland state budget. Owned by Finance Department.")

        # After discovering K is percentage:
        excel_update_skill(path, "Key Columns",
            "- `Utilisation_Pct` (col K): percentage stored as 0-1 decimal. "
            "Multiply by 100 for display. NEVER sum this column.")

        # After a successful chart, linking to the saved query:
        excel_update_skill(path, "Useful Analyses",
            "- **Utilisation barh**: shows dept utilisation clearly [query: 00dd8c59]",
            query_key="00dd8c59")
    """
    from workbook_skill import update_skill, SECTIONS

    p = Path(path).resolve()
    if not p.exists():
        return f"ERROR: File not found: {path}"
    if mode not in ("append","replace","prepend"):
        return f"ERROR: mode must be append|replace|prepend"

    try:
        result = update_skill(str(p), section, content, mode, query_key)
        if not result["ok"]:
            return (f"ERROR: {result.get('error')}\n"
                    f"Valid sections: {SECTIONS}")
        return (
            f"✅ Skill document updated.\n"
            f"- **Section:** {result['section']}\n"
            f"- **Mode:** {result['mode']}\n"
            f"- **Path:** `{result['path']}`\n\n"
            f"Call `excel_read_skill` to see the full document."
        )
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


# ─────────────────────────────────────────────────────────────────────────────
# CORPUS BUILDER TOOLS
# ─────────────────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_extract_to_kb(
    entries_json: str,
    source_label: str = "",
    skip_invalid: bool = True,
) -> str:
    """
    Batch-insert knowledge entries extracted from a PDF or book into the
    knowledge base. The LLM reads the document, extracts entries in the
    correct format, then calls this tool to persist them.

    This is the primary tool for corpus building. Workflow:
      1. LLM reads PDF/book content (via context window or file path)
      2. LLM extracts structured entries using the format below
      3. LLM calls excel_extract_to_kb(entries_json=...) with all entries
      4. MCP validates, deduplicates, and inserts
      5. Entries immediately searchable via excel_search_docs

    ENTRY FORMAT (JSON array):
    [
      {
        "title":       "Short descriptive title (10-120 chars)",
        "problem":     "One sentence: what problem this solves",
        "code":        "Complete Python/pandas/SQL using read_sheet() helpers",
        "category":    "pandas|sql|analysis|visualisation|excel|dedup",
        "subcategory": "groupby|pivot|filter|window|date|chart|...",
        "tags":        ["keyword1", "keyword2"],
        "notes":       "Gotchas, edge cases (optional)",
        "source":      "Book Title — Author, Chapter X",
        "difficulty":  "beginner|intermediate|advanced"
      }
    ]

    CODE RULES (enforced by validator):
      - Must use pre-loaded helpers: read_sheet(), save_chart(), pd., df.
      - Use placeholder names: "Category_Column", "Numeric_Col"
      - Must include print() or save_chart() for visible output
      - Do not use pd.read_excel() directly — use read_sheet()

    To get the extraction prompt for a document, call:
      excel_get_extraction_prompt(document_title="Python for Data Analysis")

    Args:
        entries_json:  JSON array of entry dicts (string)
        source_label:  Optional label applied to entries without a source
                       e.g. "Python for Data Analysis — McKinney, Ch.10"
        skip_invalid:  If True (default), inserts valid entries and reports failures.
                       If False, aborts if any entry is invalid.
    """
    from corpus_builder import batch_insert_validated, render_insert_report

    db_path = str(Path(__file__).parent / "excel_knowledge.db")

    try:
        raw = json.loads(entries_json)
        if isinstance(raw, dict) and "entries" in raw:
            entries = raw["entries"]
        elif isinstance(raw, list):
            entries = raw
        else:
            return "ERROR: entries_json must be a JSON array or {\"entries\": [...]}"
    except json.JSONDecodeError as e:
        return f"ERROR: Invalid JSON: {e}"

    try:
        result = batch_insert_validated(
            entries,
            db_path      = db_path,
            source_label = source_label,
            skip_invalid = skip_invalid,
        )
        return render_insert_report(result, context=source_label)
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_import_kb_file(
    file_path: str,
    source_label: str = "",
) -> str:
    """
    Import knowledge entries from a JSON or Markdown file into the knowledge base.

    Accepts:
      - JSON file: list of entry dicts, or {"entries": [...]}
      - Markdown file: structured entries separated by --- (see example_doc.md format)

    Use this for bulk import after a coding agent has pre-processed a book:
      1. Coding agent reads book
      2. Writes entries to a .json or .md file
      3. Call excel_import_kb_file(file_path=...) to insert all

    Args:
        file_path:    Path to .json or .md file containing entries
        source_label: Optional source label for entries without one
    """
    from corpus_builder import (import_from_json_file,
                                    import_from_markdown_file,
                                    render_insert_report)

    db_path = str(Path(__file__).parent / "excel_knowledge.db")
    p = Path(file_path)

    if not p.exists():
        return f"ERROR: File not found: {file_path}"

    try:
        if p.suffix.lower() == ".json":
            result = import_from_json_file(str(p), db_path, source_label)
        elif p.suffix.lower() in (".md", ".markdown"):
            result = import_from_markdown_file(str(p), db_path, source_label)
        else:
            return f"ERROR: Unsupported file type '{p.suffix}'. Use .json or .md"

        return render_insert_report(result, context=str(p.name))
    except Exception:
        return f"ERROR:\n```\n{traceback.format_exc()}\n```"


@mcp.tool()
def excel_get_extraction_prompt(
    document_title: str = "",
    chapter: str = "",
) -> str:
    """
    Get the extraction prompt to use when reading a PDF or book for corpus building.

    The LLM should:
      1. Call this to get the extraction prompt
      2. Read the document content
      3. Apply the prompt to extract entries
      4. Call excel_extract_to_kb with the extracted entries

    Args:
        document_title: Title of the book/document being processed
                        e.g. "Python for Data Analysis"
        chapter:        Chapter or section being processed
                        e.g. "Chapter 10 — GroupBy Mechanics"
    """
    from corpus_builder import get_extraction_prompt
    prompt = get_extraction_prompt(document_title, chapter)
    return (
        f"## Extraction Prompt for: {document_title or 'Document'}\n\n"
        f"Apply this prompt to the document content to extract knowledge entries:\n\n"
        f"```\n{prompt.strip()}\n```\n\n"
        f"After extraction, call:\n"
        f"```\nexcel_extract_to_kb(\n"
        f"    entries_json='[...your extracted entries...]',\n"
        f"    source_label='{document_title}{(', ' + chapter) if chapter else ''}'\n"
        f")\n```"
    )


if __name__ == "__main__":
    import logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    logger = logging.getLogger(__name__)
    
    logger.info("Starting Excel MCP Server v2 on port 6699")
    logger.info("MCP endpoint: http://localhost:6699/mcp")
    
    mcp.run(transport="streamable-http")