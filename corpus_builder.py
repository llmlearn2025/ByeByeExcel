"""
corpus_builder.py — PDF and document extraction for knowledge base population.

PURPOSE
───────
Enables an LLM to read a PDF or document and directly populate the Excel MCP v2
knowledge base. The LLM does the semantic understanding; this module handles the
mechanical parts: validation, deduplication, batch insert, progress tracking.

TWO MODES
─────────
Mode 1: LLM extracts → calls excel_extract_to_kb
  The LLM reads a PDF/book (via its own context window or passed content),
  extracts structured entries, and calls this tool to batch-insert them.
  The tool validates format, deduplicates, and returns a report.

Mode 2: Structured JSON file → calls excel_import_kb_file
  For bulk import from a pre-extracted JSON or markdown file.
  Used when a coding agent has pre-processed a book into the standard format.

EXTRACTION CONTRACT FOR THE LLM
────────────────────────────────
When extracting from a book, the LLM must produce entries in this exact format:

[
  {
    "title":       "Short descriptive title (what this does)",
    "problem":     "Plain English: what problem this solves, in one sentence",
    "code":        "Complete runnable Python/SQL code using read_sheet() helpers",
    "category":    "pandas|sql|analysis|visualisation|excel|dedup",
    "subcategory": "groupby|pivot|filter|window|date|merge|chart|...",
    "tags":        ["keyword1", "keyword2", "keyword3"],
    "notes":       "Gotchas, when not to use, edge cases (optional)",
    "source":      "Book Title — Author, Chapter X, Page Y",
    "difficulty":  "beginner|intermediate|advanced"
  }
]

QUALITY CHECKS (run before insert)
────────────────────────────────────
  - title must be ≥ 10 characters and ≤ 120 characters
  - problem must be ≥ 15 characters
  - code must be non-empty and contain at least one of:
    read_sheet, pd., df., plt., sqlite3, SQL
  - category must be in valid list
  - difficulty must be in valid list
  - tags must be a list with at least one entry
"""

import json, hashlib, datetime
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

VALID_CATEGORIES   = {"pandas","sql","analysis","visualisation","excel","dedup"}
VALID_DIFFICULTIES = {"beginner","intermediate","advanced"}

# At least one of these must appear in the code
CODE_SIGNALS = ["read_sheet","pd.","df.","plt.","sqlite3","SQL",
                "import","def ","groupby","merge","pivot","plot"]


def validate_entry(entry: dict) -> dict:
    """
    Validate one knowledge entry.
    Returns {"ok": bool, "errors": [str], "warnings": [str], "entry": dict}
    """
    errors   = []
    warnings = []

    title   = str(entry.get("title","")).strip()
    problem = str(entry.get("problem","")).strip()
    code    = str(entry.get("code","")).strip()
    cat     = str(entry.get("category","")).strip().lower()
    diff    = str(entry.get("difficulty","intermediate")).strip().lower()
    tags    = entry.get("tags",[])
    source  = str(entry.get("source","")).strip()

    # Required field checks
    if len(title) < 10:
        errors.append(f"title too short ({len(title)} chars, min 10): '{title[:40]}'")
    if len(title) > 120:
        errors.append(f"title too long ({len(title)} chars, max 120)")
    if len(problem) < 15:
        errors.append(f"problem too short ({len(problem)} chars, min 15)")
    if not code:
        errors.append("code is empty")
    elif not any(sig in code for sig in CODE_SIGNALS):
        warnings.append(
            f"code may not be runnable — none of {CODE_SIGNALS[:4]} found in code")
    if cat not in VALID_CATEGORIES:
        errors.append(f"category '{cat}' invalid. Valid: {sorted(VALID_CATEGORIES)}")
    if diff not in VALID_DIFFICULTIES:
        warnings.append(f"difficulty '{diff}' invalid, defaulting to 'intermediate'")
        entry["difficulty"] = "intermediate"
    if not isinstance(tags, list) or len(tags) == 0:
        warnings.append("no tags provided — searchability reduced")
        entry["tags"] = []
    if not source:
        warnings.append("no source provided")

    # Auto-fix: lowercase category
    if cat in VALID_CATEGORIES:
        entry["category"] = cat

    # Auto-fix: ensure tags is a list of strings
    if isinstance(tags, str):
        entry["tags"] = [t.strip() for t in tags.split(",") if t.strip()]

    return {
        "ok":       len(errors) == 0,
        "errors":   errors,
        "warnings": warnings,
        "entry":    entry,
    }


def validate_batch(entries: list) -> dict:
    """
    Validate a list of entries.
    Returns summary with valid/invalid counts and per-entry results.
    """
    results = [validate_entry(e) for e in entries]
    valid   = [r for r in results if r["ok"]]
    invalid = [r for r in results if not r["ok"]]
    warned  = [r for r in results if r["warnings"]]
    return {
        "total":   len(entries),
        "valid":   len(valid),
        "invalid": len(invalid),
        "warned":  len(warned),
        "results": results,
    }


# ─────────────────────────────────────────────────────────────────────────────
# BATCH INSERT
# ─────────────────────────────────────────────────────────────────────────────

def batch_insert_validated(
    entries: list,
    db_path: str,
    source_label: str = "",
    skip_invalid: bool = True,
) -> dict:
    """
    Validate then bulk-insert entries into the knowledge base.
    Returns a detailed report.

    skip_invalid: if True, inserts valid entries and reports failures.
                  if False, aborts if any entry is invalid.
    """
    from knowledge import init_knowledge_db, bulk_insert

    init_knowledge_db(db_path)
    validation = validate_batch(entries)

    if not skip_invalid and validation["invalid"] > 0:
        return {
            "ok":      False,
            "error":   f"{validation['invalid']} invalid entries. Set skip_invalid=True to insert valid ones.",
            "details": validation,
        }

    valid_entries = [r["entry"] for r in validation["results"] if r["ok"]]

    if source_label:
        for e in valid_entries:
            if not e.get("source"):
                e["source"] = source_label

    insert_result = bulk_insert(valid_entries, db_path)

    return {
        "ok":           True,
        "validated":    validation["total"],
        "valid":        validation["valid"],
        "invalid":      validation["invalid"],
        "warned":       validation["warned"],
        "inserted":     insert_result["inserted"],
        "skipped_dup":  insert_result["skipped"],
        "db_path":      db_path,
        "invalid_details": [
            {"title": r["entry"].get("title","?"), "errors": r["errors"]}
            for r in validation["results"] if not r["ok"]
        ],
        "warning_details": [
            {"title": r["entry"].get("title","?"), "warnings": r["warnings"]}
            for r in validation["results"] if r["warnings"]
        ],
    }


# ─────────────────────────────────────────────────────────────────────────────
# JSON FILE IMPORT
# ─────────────────────────────────────────────────────────────────────────────

def import_from_json_file(
    json_path: str,
    db_path: str,
    source_label: str = "",
) -> dict:
    """
    Import entries from a JSON file (list of entry dicts).
    Validates each entry before inserting.
    """
    p = Path(json_path)
    if not p.exists():
        return {"ok": False, "error": f"File not found: {json_path}"}

    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        return {"ok": False, "error": f"JSON parse error: {e}"}

    if isinstance(raw, dict) and "entries" in raw:
        entries = raw["entries"]
    elif isinstance(raw, list):
        entries = raw
    else:
        return {"ok": False, "error": "JSON must be a list or {\"entries\": [...]}"}

    label = source_label or p.stem
    return batch_insert_validated(entries, db_path, source_label=label)


# ─────────────────────────────────────────────────────────────────────────────
# MARKDOWN FILE IMPORT (from structured extraction)
# ─────────────────────────────────────────────────────────────────────────────

def import_from_markdown_file(
    md_path: str,
    db_path: str,
    source_label: str = "",
) -> dict:
    """
    Parse a markdown file of knowledge entries and insert them.

    Expected format (each entry separated by ---):

    ## Title of the entry
    **Problem:** What this solves
    **Category:** pandas/groupby
    **Tags:** tag1, tag2, tag3
    **Difficulty:** intermediate
    **Source:** Book Name, Ch.X
    ```python
    code here
    ```
    **Notes:** gotcha or edge case
    ---
    """
    import re
    p = Path(md_path)
    if not p.exists():
        return {"ok": False, "error": f"File not found: {md_path}"}

    content = p.read_text(encoding="utf-8")
    entries = _parse_md_entries(content)

    if not entries:
        return {"ok": False, "error": "No valid entries found in markdown file"}

    label = source_label or p.stem
    return batch_insert_validated(entries, db_path, source_label=label)


def _parse_md_entries(content: str) -> list:
    import re
    entries = []
    blocks = re.split(r"\n---\n", content)
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        try:
            title_m = re.search(r"^## (.+)", block, re.M)
            if not title_m:
                continue
            title = title_m.group(1).strip()

            def _field(name):
                m = re.search(rf"\*\*{name}:\*\*\s*(.+)", block)
                return m.group(1).strip() if m else ""

            problem  = _field("Problem")
            cat_str  = _field("Category")
            tags_str = _field("Tags")
            diff     = _field("Difficulty") or "intermediate"
            source   = _field("Source")
            notes    = _field("Notes")

            code_m = re.search(r"```(?:python|sql|)\n(.+?)```", block, re.DOTALL)
            code   = code_m.group(1).strip() if code_m else ""

            cat_parts   = cat_str.split("/")
            category    = cat_parts[0].strip().lower()
            subcategory = cat_parts[1].strip() if len(cat_parts) > 1 else ""
            tags = [t.strip() for t in tags_str.split(",") if t.strip()]

            if title and problem and code:
                entries.append({
                    "category":    category or "pandas",
                    "subcategory": subcategory,
                    "title":       title,
                    "problem":     problem,
                    "tags":        tags,
                    "code":        code,
                    "notes":       notes,
                    "source":      source,
                    "difficulty":  diff,
                })
        except Exception:
            continue
    return entries


# ─────────────────────────────────────────────────────────────────────────────
# RENDER REPORT
# ─────────────────────────────────────────────────────────────────────────────

def render_insert_report(result: dict, context: str = "") -> str:
    """Human-readable markdown report of a batch insert operation."""
    L = []; a = L.append

    if not result.get("ok"):
        a(f"## ❌ Insert Failed\n{result.get('error','unknown error')}")
        return "\n".join(L)

    a(f"## Knowledge Base Import Complete")
    if context:
        a(f"*Source: {context}*\n")

    a(f"| Metric | Count |")
    a(f"|---|---|")
    a(f"| Entries validated | {result['validated']:,} |")
    a(f"| Valid entries | {result['valid']:,} |")
    a(f"| Invalid (skipped) | {result['invalid']:,} |")
    a(f"| Inserted (new) | {result['inserted']:,} |")
    a(f"| Skipped (duplicate) | {result['skipped_dup']:,} |")

    if result.get("invalid_details"):
        a(f"\n### ❌ Invalid entries (not inserted)")
        for d in result["invalid_details"][:10]:
            a(f"- **{d['title'][:60]}**: {'; '.join(d['errors'])}")
        if len(result["invalid_details"]) > 10:
            a(f"- *... and {len(result['invalid_details'])-10} more*")

    if result.get("warning_details"):
        a(f"\n### ⚠️ Warnings (inserted with caution)")
        for d in result["warning_details"][:5]:
            a(f"- **{d['title'][:60]}**: {'; '.join(d['warnings'])}")

    if result["inserted"] > 0:
        a(f"\n✅ **{result['inserted']} new entries** immediately searchable via `excel_search_docs`.")

    return "\n".join(L)


# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTION PROMPT TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────

EXTRACTION_PROMPT = """
You are extracting knowledge entries from a document for the Excel MCP v2 knowledge base.

For each distinct technique, pattern, or code example in the document, produce one JSON entry.

REQUIRED FORMAT (list of objects):
[
  {
    "title":       "Short action-oriented title (10-120 chars)",
    "problem":     "One sentence: what specific problem does this solve?",
    "code":        "Complete runnable Python/pandas/SQL code. Use these pre-loaded helpers: read_sheet(sheet_name), read_all_sheets(), read_sheet_chunked(sheet, chunksize), save_chart(name), save_new_workbook(df, sheet_name), AI_WORKBOOK, OUTPUT_DIR. Use placeholder column names like 'Category_Column', 'Numeric_Col'.",
    "category":    "ONE OF: pandas | sql | analysis | visualisation | excel | dedup",
    "subcategory": "e.g. groupby | pivot | filter | window | date | merge | chart | formula",
    "tags":        ["keyword1", "keyword2", "keyword3"],
    "notes":       "Optional: gotchas, edge cases, when NOT to use",
    "source":      "{DOCUMENT_TITLE}, {CHAPTER_OR_SECTION}",
    "difficulty":  "ONE OF: beginner | intermediate | advanced"
  }
]

QUALITY RULES:
1. Code must be runnable — use the pre-loaded helpers, not bare pd.read_excel()
2. Use placeholder column names so the LLM knows to substitute from the graph
3. Each entry must solve ONE specific problem, not explain a concept
4. Include print() or save_chart() so there is visible output
5. Add the percentage column warning if the code touches % columns
6. Do NOT include entries that just describe what a function does

OUTPUT: Return ONLY the JSON array. No preamble. No explanation. No markdown fences.
"""

def get_extraction_prompt(document_title: str = "", chapter: str = "") -> str:
    """Return the extraction prompt with document context filled in."""
    source = document_title
    if chapter:
        source += f", {chapter}"
    return EXTRACTION_PROMPT.replace("{DOCUMENT_TITLE}", document_title or "Document") \
                             .replace("{CHAPTER_OR_SECTION}", chapter or "")
