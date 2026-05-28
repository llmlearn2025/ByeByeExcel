"""
workbook_skill.py — Per-workbook living skill document.

PURPOSE
───────
A workbook skill document is the LLM's accumulated analytical understanding
of ONE specific workbook, written in structured prose + rules. It is distinct
from both the graph (structural metadata) and queries.json (executable code).

  graph.json          → "what columns exist, what types, what formats"
  queries.json        → "here is code that worked on this workbook"
  workbook_skill.md   → "here is how to THINK about this workbook"

The skill document captures things the graph cannot:
  - Domain context: "This is the Nagaland state budget tracking file FY 2025-26"
  - Analytical rules: "Always group by Department before District"
  - Warnings: "Column K (Utilisation_Pct) breaks if you write to it directly"
  - Validated combinations: "The most useful chart is barh of K by B"
  - User preferences: "User prefers ₹ Crore display, not raw rupees"
  - Data quality notes: "Rows 166-167 are known duplicates, excluded from analysis"
  - Linkage: "See queries.json key 'abc123' for working utilisation chart code"

LIFECYCLE
─────────
  excel_init       → checks for existing skill MD, includes it in init output
  LLM works           → discovers something important about the workbook
  excel_update_skill  → LLM writes the insight to the skill document
  next session        → excel_read_skill returns full document to LLM context

STRUCTURE
─────────
The skill MD has named sections. Each section is append-only unless the LLM
explicitly replaces it. Sections:

  ## Context
  Domain description, fiscal year, owning department, purpose of the file.

  ## Key Columns
  Which columns matter most, how to interpret them, gotchas.

  ## Analytical Rules
  Validated patterns: "always do X before Y", "never sum column K".

  ## Useful Analyses
  Named analyses that have worked well, with a link to their query key.

  ## Warnings
  Things that break, data quality issues, cells to avoid.

  ## User Preferences
  Display format preferences, grouping preferences, chart preferences.

  ## History
  Timestamped log of significant findings and changes.

LINK TO QUERIES.JSON
────────────────────
The skill document and queries.json are linked:
  - Skill "Useful Analyses" entries reference queries.json keys
  - When excel_search_docs finds a query memory match, it also checks
    the skill document for context about that query
  - excel_read_skill returns both the skill MD and the referenced query codes
"""

import re, json, datetime
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# PATH HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _skill_path(ai_path: str) -> Path:
    p = Path(ai_path).resolve()
    return p.parent / (p.stem.replace("_ai_workbook","") + "_skill.md")

def _queries_path(ai_path: str) -> Path:
    p = Path(ai_path).resolve()
    return p.parent / (p.stem.replace("_ai_workbook","") + "_queries.json")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────

SECTIONS = [
    "Context",
    "Key Columns",
    "Analytical Rules",
    "Useful Analyses",
    "Warnings",
    "User Preferences",
    "History",
]

_SECTION_GUIDANCE = {
    "Context": (
        "Domain context: what this workbook is, who owns it, what period it covers, "
        "what department uses it. 2-5 sentences."
    ),
    "Key Columns": (
        "One bullet per important column: name, what it means, how to interpret it, "
        "any gotchas. Include format notes (e.g. percentage stored as 0-1 decimal)."
    ),
    "Analytical Rules": (
        "Validated rules the LLM should always follow for this workbook. "
        "E.g. 'Always sort by Department before grouping', "
        "'Never write to the Summary sheet — it is formula-driven'."
    ),
    "Useful Analyses": (
        "Named analyses that have worked well. Each entry: analysis name, "
        "what it shows, query key from queries.json (if saved). "
        "Format: `- **Name**: description [query: {key}]`"
    ),
    "Warnings": (
        "Things that break, known data quality issues, cells or ranges to avoid. "
        "E.g. 'Rows 166-167 are duplicate test entries'. "
        "E.g. 'Column J formula breaks if you modify column F'."
    ),
    "User Preferences": (
        "Display and presentation preferences. "
        "E.g. 'User prefers ₹ Crore (not raw rupees)', "
        "'Horizontal bar charts preferred over vertical', "
        "'Sort descending by amount in all tables'."
    ),
    "History": (
        "Timestamped log. Auto-managed. Do not edit manually."
    ),
}


# ─────────────────────────────────────────────────────────────────────────────
# READ
# ─────────────────────────────────────────────────────────────────────────────

def read_skill(ai_path: str) -> dict:
    """
    Read the workbook skill document.
    Returns:
    {
        "exists":    bool,
        "path":      str,
        "content":   str (full markdown),
        "sections":  {section_name: content_str},
        "query_refs": [query_keys referenced in Useful Analyses],
        "linked_queries": [{key, description, code, ...}],
    }
    """
    path = _skill_path(ai_path)
    result = {
        "exists":         path.exists(),
        "path":           str(path),
        "content":        "",
        "sections":       {},
        "query_refs":     [],
        "linked_queries": [],
    }

    if not path.exists():
        return result

    content = path.read_text(encoding="utf-8")
    result["content"] = content
    result["sections"] = _parse_sections(content)

    # Extract query keys referenced in Useful Analyses
    useful = result["sections"].get("Useful Analyses","")
    keys = re.findall(r'\[query:\s*([a-f0-9]{8,16})\]', useful)
    result["query_refs"] = keys

    # Load referenced queries
    if keys:
        qpath = _queries_path(ai_path)
        if qpath.exists():
            try:
                mem = json.loads(qpath.read_text())
                all_q = {q["key"]: q for q in mem.get("queries",[])}
                result["linked_queries"] = [
                    all_q[k] for k in keys if k in all_q
                ]
            except Exception:
                pass

    return result


def _parse_sections(content: str) -> dict:
    """Split markdown into {section_name: content} dict."""
    sections = {}
    current  = None
    buf      = []
    for line in content.splitlines():
        m = re.match(r"^## (.+)$", line)
        if m:
            if current and buf:
                sections[current] = "\n".join(buf).strip()
            current = m.group(1).strip()
            buf = []
        elif current is not None:
            buf.append(line)
    if current and buf:
        sections[current] = "\n".join(buf).strip()
    return sections


# ─────────────────────────────────────────────────────────────────────────────
# WRITE
# ─────────────────────────────────────────────────────────────────────────────

def init_skill(ai_path: str, context: str = "") -> str:
    """
    Create a new skill document for a workbook.
    Only creates if it doesn't already exist.
    Returns the path.
    """
    path = _skill_path(ai_path)
    if path.exists():
        return str(path)

    p = Path(ai_path)
    stem = p.stem.replace("_ai_workbook","")
    now  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        f"# Workbook Skill: {stem}",
        f"*Auto-generated by Excel MCP v2 — edit via excel_update_skill*",
        f"*Linked queries: `{stem}_queries.json`*",
        "",
    ]
    for section in SECTIONS:
        lines.append(f"## {section}")
        if section == "Context" and context:
            lines.append(context)
        elif section == "History":
            lines.append(f"- {now}: Skill document created")
        else:
            lines.append(f"*{_SECTION_GUIDANCE[section]}*")
        lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    return str(path)


def update_skill(
    ai_path: str,
    section: str,
    content: str,
    mode: str = "append",
    query_key: str = "",
) -> dict:
    """
    Update a named section of the workbook skill document.

    section:   One of: Context | Key Columns | Analytical Rules |
               Useful Analyses | Warnings | User Preferences | History
    content:   Text to add or replace. For "Useful Analyses", include
               [query: {key}] references to queries.json entries.
    mode:      "append"  → add to existing section content (default)
               "replace" → replace section content entirely
               "prepend" → add before existing content
    query_key: If provided, creates a cross-reference link in the History.

    Returns: {"path": str, "section": str, "mode": str, "ok": bool}
    """
    path = _skill_path(ai_path)

    # Auto-init if missing
    if not path.exists():
        init_skill(ai_path)

    current_content = path.read_text(encoding="utf-8")
    sections = _parse_sections(current_content)

    # Normalise section name (case-insensitive partial match)
    matched = _match_section(section)
    if not matched:
        return {
            "ok": False,
            "error": f"Unknown section '{section}'. Valid: {SECTIONS}",
        }

    # Prepare new section content
    existing = sections.get(matched, "").strip()
    # Strip the placeholder guidance text if it's still there
    if existing.startswith("*") and existing.endswith("*"):
        existing = ""

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    if mode == "replace":
        new_section_content = content.strip()
    elif mode == "prepend":
        new_section_content = content.strip() + ("\n\n" + existing if existing else "")
    else:  # append
        new_section_content = (existing + "\n\n" + content.strip()
                               if existing else content.strip())

    # Auto-add to History
    hist_entry = f"- {now}: Updated [{matched}]"
    if query_key:
        hist_entry += f" [query: {query_key}]"
    hist_existing = sections.get("History","").strip()
    new_history = hist_entry + ("\n" + hist_existing if hist_existing else "")

    # Rebuild full document
    sections[matched]  = new_section_content
    sections["History"] = new_history
    new_content = _rebuild_document(ai_path, sections)
    path.write_text(new_content, encoding="utf-8")

    return {
        "ok":      True,
        "path":    str(path),
        "section": matched,
        "mode":    mode,
    }


def _match_section(name: str) -> Optional[str]:
    """Case-insensitive partial match against SECTIONS list."""
    nl = name.lower().strip()
    for s in SECTIONS:
        if nl == s.lower() or nl in s.lower():
            return s
    return None


def _rebuild_document(ai_path: str, sections: dict) -> str:
    """Rebuild full markdown from sections dict, preserving order."""
    p = Path(ai_path)
    stem = p.stem.replace("_ai_workbook","")
    lines = [
        f"# Workbook Skill: {stem}",
        f"*Auto-generated by Excel MCP v2 — edit via excel_update_skill*",
        f"*Linked queries: `{stem}_queries.json`*",
        "",
    ]
    for section in SECTIONS:
        lines.append(f"## {section}")
        sc = sections.get(section,"").strip()
        if sc:
            lines.append(sc)
        else:
            lines.append(f"*{_SECTION_GUIDANCE[section]}*")
        lines.append("")
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# RENDER — for LLM context
# ─────────────────────────────────────────────────────────────────────────────

def render_skill_for_llm(ai_path: str, include_queries: bool = True) -> str:
    """
    Render the skill document for inclusion in LLM context.
    Includes referenced query code inline for immediate use.
    """
    skill = read_skill(ai_path)

    if not skill["exists"]:
        return (
            f"*No skill document for this workbook yet.*\n"
            f"Call `excel_update_skill` to start building it.\n"
            f"Path would be: `{skill['path']}`"
        )

    L = []; a = L.append
    a(skill["content"])

    if include_queries and skill["linked_queries"]:
        a("\n---\n## Referenced Query Code\n")
        a("*Code linked from Useful Analyses section:*\n")
        for q in skill["linked_queries"]:
            a(f"### {q['description']}")
            a(f"- Key: `{q['key']}`  |  Run count: {q.get('execution_count',1)}")
            if q.get("sheet"):
                a(f"- Sheet: `{q['sheet']}`")
            if q.get("result_summary"):
                a(f"- Last result: {q['result_summary']}")
            a(f"```python\n{q['code'].strip()}\n```\n")

    return "\n".join(L)


def skill_exists(ai_path: str) -> bool:
    return _skill_path(ai_path).exists()


def get_skill_summary(ai_path: str) -> dict:
    """Compact summary for inclusion in graph init output."""
    path = _skill_path(ai_path)
    if not path.exists():
        return {"exists": False}
    skill = read_skill(ai_path)
    ctx   = skill["sections"].get("Context","").strip()
    rules = skill["sections"].get("Analytical Rules","").strip()
    warn  = skill["sections"].get("Warnings","").strip()
    return {
        "exists":         True,
        "path":           str(path),
        "context_preview":ctx[:200] if ctx else "",
        "rule_count":     len([l for l in rules.splitlines() if l.strip().startswith("-")]),
        "warning_count":  len([l for l in warn.splitlines() if l.strip().startswith("-")]),
        "linked_queries": len(skill["query_refs"]),
    }
