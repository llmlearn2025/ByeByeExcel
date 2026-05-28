# Excel MCP v2 — Design Document

> This document is written for a coding agent running autonomous tests.
> Read it fully before touching any code. It tells you what exists, what each
> piece does, how they connect, and what the test script verifies.

---

## 1. What This Is

Excel MCP v2 is a Model Context Protocol server that gives any LLM agent
(Claude Code, OpenCode, LM Studio, Cursor) the ability to work correctly
with Excel files — including files with 11 lakh (1.1 million) rows.

**Core design principle:**
- MCP = execution engine (runs code, manages files, searches knowledge)
- LLM = intelligence (decides what to do, chooses strategies, interprets results)

The MCP never hardcodes domain logic. It exposes primitives. The LLM decides
which primitives to combine and how.

---

## 2. File Map

```
excel_mcp/
│
├── server.py          ← ENTRY POINT. All 19 MCP tools live here.
│                            Start with: python server.py
│
├── graph.py           ← Workbook inspection (2 passes)
│                            Pass 1: openpyxl read_only=True — structure, formulas
│                            Pass 2: openpyxl read_only=False rows 1-2 — formatting
│
├── storage.py         ← Transparent pandas/SQLite switch
│                            < 50K rows → pandas in-memory
│                            ≥ 50K rows → SQLite chunked
│
├── normalisers.py     ← Pluggable text normalisation for fuzzy dedup
│                            LLM picks normaliser; MCP executes
│                            Available: naga_names, indian_names, generic_text,
│                                       address, numeric_range, date_approx, code_id
││
├── audit.py           ← Data quality checks (8 issue types)
│                            FORMULA_OVERRIDE, LOGIC_VIOLATION, TYPE_VIOLATION,
│                            FORMAT_VIOLATION, RANGE_VIOLATION, MISSING_REQUIRED,
│                            DUPLICATE_ROW, CONSISTENCY
│
├── ingest.py          ← Universal column mapper + Excel→SQLite ingestion
│                            Auto-maps any header to canonical _c_* fields
│                            Streams in chunks — safe for 11L rows
│
├── knowledge.py       ← Doc search + workbook query memory
│                            Subsystem 1: SQLite FTS5 knowledge base
│                            Subsystem 2: Per-workbook _queries.json
│
├── workbook_skill.py  ← Per-workbook skill document (_skill.md)
│                            Stores conceptual understanding, not just code
│                            Linked to _queries.json via [query: key] refs
│
├── corpus_builder.py  ← PDF/book → knowledge base pipeline
│                            Validates, deduplicates, bulk-inserts entries
│                            Entry format validation before insert
│
├── rich_output.py     ← Smart output tiering
│                            ≤40 rows → markdown inline
│                            41-500 → HTML + base64 thumbnail
│                            >500 → Excel download + chart
│
├── seed_corpus.py     ← 29 seed knowledge entries across 6 categories
│                            Run once: python seed_corpus.py
│
├── excel_knowledge.db    ← SQLite FTS5 knowledge base (auto-created by seed)
│
├── SKILL.md           ← Instructions for LLM agents using this MCP
├── example_doc.md        ← Instructions for corpus-building coding agents
│
└── (v1 preserved, not deleted)
    excel_mcp/server.py         ← V1 server (intact)
    beneficiary_pipeline/        ← V1 beneficiary pipeline (intact)
```

**Per-workbook files (created alongside the .xlsx):**
```
{stem}_ai_workbook.xlsx   ← Working copy. Original never touched after init.
{stem}_graph.json         ← Structural + format graph
{stem}_skill.md           ← Conceptual skill document (7 sections)
{stem}_queries.json       ← Executable query memory
{stem}_analysis/          ← Charts, exports, audit reports
{stem}.db                 ← SQLite (only created for large files via excel_ingest_large)
```

---

## 3. The 19 MCP Tools

### Workspace Management
| Tool | Purpose |
|---|---|
| `excel_init(path)` | Copy → graph → format scan → create skill doc → check query memory |
| `excel_inspect(path)` | Refresh graph after edits |

### Analysis
| Tool | Purpose |
|---|---|
| `excel_execute(path, task)` | Plain-English task execution |
| `excel_run_code(path, code)` | Custom Python with pre-loaded helpers |
| `excel_audit(path)` | Data quality (8 issue types) — on demand only |
| `excel_apply_fixes(path, fixes_json)` | Write LLM-proposed fixes |

### Large File Pipeline
| Tool | Purpose |
|---|---|
| `excel_ingest_large(path, db_path)` | Excel → SQLite (for 11L+ row files) |
| `excel_bogus_detect(path)` | Fraud/anomaly detection — on demand only |

### Knowledge Base
| Tool | Purpose |
|---|---|
| `excel_search_docs(query, path, ...)` | Search docs + workbook memory |
| `excel_save_query(path, description, code, ...)` | Save working code to workbook memory |
| `excel_add_doc_entry(title, problem, code, ...)` | Add one entry to global KB |
| `excel_kb_stats()` | Show KB statistics |

### Corpus Builder
| Tool | Purpose |
|---|---|
| `excel_get_extraction_prompt(doc_title, chapter)` | Get prompt for PDF extraction |
| `excel_extract_to_kb(entries_json, source_label)` | Batch-insert LLM-extracted entries |
| `excel_import_kb_file(file_path)` | Import from .json or .md file |

### Workbook Skill
| Tool | Purpose |
|---|---|
| `excel_read_skill(path)` | Read skill doc + linked query code |
| `excel_update_skill(path, section, content, ...)` | Write to skill doc |

---

## 4. Data Flows

### Normal workflow (small file)
```
excel_init("Budget.xlsx")
    → copies to Budget_ai_workbook.xlsx
    → builds Budget_graph.json (structure + formatting)
    → creates Budget_skill.md (empty)
    → returns graph markdown + skill summary

excel_search_docs("utilisation chart by department", path=ai_path)
    → searches FTS5 knowledge base (BM25 ranked)
    → searches workbook query memory (keyword overlap)
    → returns workbook memory first, then generic docs

excel_run_code(ai_path, code)  ← code adapted from search results
    → executes in sandbox with pre-loaded helpers
    → returns output text + embedded base64 chart

excel_save_query(ai_path, description, code, ...)
    → saves to Budget_queries.json
    → deduplicates by semantic key

excel_update_skill(ai_path, "Useful Analyses", "...")
    → writes to Budget_skill.md
    → timestamps in History section

(next session)
excel_init("Budget.xlsx")
    → returns: "Skill: FY 2025-26... | 3 warnings | 1 linked queries"
    → LLM calls excel_read_skill → gets full context + working code
```

### Large file workflow (11L rows)
```
excel_ingest_large("Beneficiaries_11L.xlsx", "beneficiaries.db")
    → streams in 10K chunks
    → auto-maps columns to _c_* canonical fields
    → creates indexes on all canonical columns
```

### Corpus building workflow
```
(coding agent reads a PDF chapter)

excel_get_extraction_prompt("Python for Data Analysis", "Ch.10")
    → returns structured extraction prompt

(LLM applies prompt to chapter content, produces JSON array)

excel_extract_to_kb(entries_json, source_label="McKinney Ch.10")
    → validates each entry (title length, code signals, category)
    → deduplicates by title hash
    → inserts valid entries into FTS5 table
    → returns validation report

(next search)
excel_search_docs("groupby sum by category")
    → new entry appears in results
```

---

## 5. Key Design Decisions

**Why SQLite FTS5 for knowledge search (not embeddings)?**
Zero external dependencies. BM25 ranking is native. Handles 1 lakh entries in
<5ms. Porter stemming handles word variants. Embedding models require a running
inference server, add latency, and the knowledge entries are short enough that
keyword overlap is sufficient.

**Why workbook_skill.md separate from queries.json?**
Queries JSON = executable artefacts (code + column names + execution counts).
Skill MD = conceptual understanding (domain context, rules, warnings).
The skill doc stores things that cannot be expressed as code: "The Summary
sheet is formula-driven — never edit it directly."

**Why LLM picks the normaliser (not the MCP)?**
The MCP cannot know whether "Tse Sangtam" is a Naga name, a Hindi name, or a
company name without domain knowledge. The LLM reads the graph (sample values,
fill meaning, column name) and makes an informed choice. This keeps the MCP
domain-agnostic and the LLM accountable.

**Why audit is demand only?**
Running full audit on every init adds 10-30 seconds. Users don't always want
it. The principle: init is fast (copy + graph + format scan). Heavy analysis
only when explicitly requested.

---

## 6. Dependencies

```bash
pip install fastmcp openpyxl pandas matplotlib rapidfuzz sqlite-utils
pip install oletools    # optional: VBA extraction from .xlsm files
```

Python 3.10+ required (match-case not used, but dataclasses with field() are).

No external services required. No embedding model. No Redis. No PostgreSQL.
Everything is local files + SQLite.

---

## 7. MCP Configuration

```json
{
  "mcpServers": {
    "excel-v2": {
      "command": "python",
      "args": ["/absolute/path/to/excel_mcp/server.py"]
    }
  }
}
```

SKILL.md placement:
- Claude Code: `~/.claude/skills/excel-analyst-v2/SKILL.md`
- OpenCode: `~/.opencode/skills/excel-analyst-v2/SKILL.md`

---

## 8. Known Limitations

1. **Excel row limit**: Single sheet max 1,048,576 rows. Files with 11L rows
   must span multiple sheets or use CSV. `excel_ingest_large` handles
   multi-sheet workbooks by combining all sheets.

2. **Formula values**: openpyxl `data_only=True` shows last-computed values,
   not live recalculated values. Values are stale if the file was never opened
   in Excel after formula changes. The graph notes this as `formula_override`
   when a cell has a formula but the values pass shows a hardcoded number.

3. **VBA extraction**: Requires `oletools` for full extraction. Falls back to
   raw string scanning without it (partial — may miss some modules).

4. **Date parsing**: Excel serial dates (integers like 46000) require manual
   conversion. `graph` detects `format_meaning=date` but pandas reads them
   as integers. Use `pd.to_datetime(col, unit='D', origin='1899-12-30')`.

5. **FTS5 tokeniser**: The porter stemmer handles English well. For Hindi or
   other Indian language column values, FTS5 falls back to character-level
   matching, which still works but with lower precision.

6. **Dedup at 11L scale**: The full dedup run takes ~50 minutes. This is a
   one-time batch process, not a per-query operation. After the first run,
   results are in SQLite and all subsequent queries are instant.


Here is a concise markdown documentation of the ingestion pipeline fixes.

---

# Architectural Fix: Dynamic Schema Merging & SQLite Ingestion

## Context & Architecture

When streaming massive, multi-sheet Excel workbooks into a single SQLite destination table, schema variations between sheets create structure conflicts. The ingestion layer handles this by extracting data in memory-safe chunks and normalizing individual rows to conform to a globally uniform schema.

---

## The Issues & Solutions

### 1. The Schema Mismatch (`no such column: *`)

* **The Problem:** The database table structure was locked down by whatever columns were present in the very first sheet/chunk. When subsequent sheets introduced unique columns (e.g., `Head_Code`), the `to_sql(..., if_exists="append")` operation crashed because the column didn't exist in the database table layout.
* **The Fix:** **Pre-compile a Master Schema.** Scan all worksheet headers *before* entering the processing loop. Every chunk is then dynamically restructured via `.reindex()` to match this master layout right before it is committed to SQLite. Missing columns safely default to `NaN`/`NULL`.

### 2. The Driver Limit Overflow (`too many SQL variables`)

* **The Problem:** Compiling wide data matrices across multiple sheets spiked the cumulative column count. When paired with `method="multi"` and a static `chunksize=1000`, Pandas attempted to pass more placeholders (`?`) than SQLite’s hardcoded query variable limit allows.
* **The Fix:** **Dynamic Batch Sizing.** Swap the static chunk ceiling for a calculation that scales based on the current width of the dataframe layout, ensuring total variables never breach safety limits:

$$\text{Safe Chunksize} = \lfloor 990 / \text{DataFrame Column Width} \rfloor$$



---

## Optimized Implementation Reference

```python
def ingest_excel_to_sqlite(excel_path: str, db_path: str, table: str = "data") -> dict:
    xl = pd.ExcelFile(excel_path)
    
    # 1. BUILD MASTER SCHEMA BEFORE THE LOOP
    master_cols = set()
    for sheet in xl.sheet_names:
        header_df = pd.read_excel(excel_path, sheet_name=sheet, nrows=0)
        orig_cols = list(header_df.columns)
        col_map = map_columns(orig_cols)
        
        master_cols.update(orig_cols)
        master_cols.update([f"_c_{c}" for c in col_map.values()])
        
    master_cols.update(["_source_sheet", "_source_row", "_row_hash"])
    master_column_list = list(master_cols)

    # 2. CHUNKED PROCESSING LOOP
    table_exists = False
    for sheet in xl.sheet_names:
        skip = 1
        while True:
            # ... [Read raw chunk data into DataFrame] ...

            # Apply clean transformations & row hashes
            # ...

            # FORCE SCHEMA ALIGNMENT
            chunk = chunk.reindex(columns=master_column_list)

            # CALCULATE SAFE DRIVER LIMIT BATCHES
            safe_db_chunksize = max(1, 990 // len(chunk.columns))

            # STREAM TO DATABASE
            chunk.to_sql(
                table, conn,
                if_exists="append" if table_exists else "replace",
                index=False, method="multi", chunksize=safe_db_chunksize,
            )
            table_exists = True

```