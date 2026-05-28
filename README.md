# ByeByeExcel

An intelligent Excel analysis and data processing system that brings LLM-powered insights directly from your spreadsheets.

## Overview

ByeByeExcel is a Model Context Protocol (MCP) server that gives LLM agents (Claude Code, OpenCode, Cursor, LM Studio) the ability to work intelligently with Excel files—including massive datasets with 1+ million rows.

**Core Philosophy:** The MCP provides execution primitives and file management, while the LLM handles intelligent decision-making, strategy selection, and result interpretation.

## Features

### Workbook Intelligence & Skill Management

- **Intelligent Workbook Analysis** - Automatic structure scanning, format detection, and column type classification
- **Per-Workbook Skill Documentation** - Create and maintain domain-specific skill documents linked to each Excel file. The system remembers your workbook's context, key columns, analytical rules, and validated patterns across sessions.
- **Dynamic Skill Updates** - Use `excel_update_skill` to add workbook-specific context, warnings, useful analyses, and user preferences that persist across sessions
- **Functional Code Library** - Link executable code templates directly to workbook skills for repeatable analysis patterns

### Knowledge Base & LLM Enhancement

- **Self-Service Knowledge Base** - Build a knowledge base from your own documents (PDFs, EPUBs) using the automated book ingestion pipeline in `actual_ingestion.py`
- **Document-to-Knowledge Pipeline** - Extract code examples, formulas, and analysis techniques from PDF books and convert them into searchable knowledge entries with automatic categorization
- **LLM-Aided Processing Enhancement** - The system intelligently searches your shared knowledge base for relevant patterns, then uses these to improve future processing decisions
- **Query Memory per Workbook** - Save successful queries directly to the workbook's memory. High-execution-count queries appear first in future searches

### Bulk Data Ingestion & Analysis

- **Bulk Excel Ingestion** - Use `actual_ingestion.py` to process multiple Excel files and extract Python examples, financial analysis patterns, and formula techniques into your knowledge base
- **Excel → SQLite Conversion** - Convert large workbooks (1M+ rows) to SQLite once for fast querying. Supports per-sheet tables or merged schema mode
- **Smart Output Tiering** - Results automatically route to the optimal format: inline Markdown tables (≤40 rows), HTML previews (40-500 rows), or Excel downloads (>500 rows)

## Architecture

```
ByeByeExcel/
├── server.py          # MCP server entry point (19 tools)
├── graph.py           # Workbook inspection (structure + formatting)
├── storage.py         # Smart data backend (pandas ↔ SQLite)
├── normalisers.py     # Pluggable text normalisation for fuzzy dedup
├── audit.py           # Data quality checks (8 issue types)
├── ingest.py          # Column mapping and Excel→SQLite ingestion
├── knowledge.py       # Search and memory management
├── corpus_builder.py  # PDF/book → knowledge base pipeline
├── seed_corpus.py     # Pre-populated knowledge entries
└── rich_output.py     # Smart output formatting
```

## Quick Start

### Installation

```bash
pip install fastmcp openpyxl pandas matplotlib rapidfuzz sqlite-utils
```

Python 3.10+ required.

### Running the Server

```bash
python server.py
```

The MCP server starts on port 6699 at `http://localhost:6699/mcp`.

## MCP Tools

The system provides **19 tools** for LLM agents:

| Category | Tool | Purpose |
|----------|------|---------|
| **Workspace** | `excel_init` | Initialize workbook (copy, graph, format scan) |
| **Analysis** | `excel_execute` | Plain-English task execution |
| **Analysis** | `excel_run_code` | Custom Python with pre-loaded helpers |
| **Analysis** | `excel_audit` | Data quality checks (8 issue types) |
| **Large Files** | `excel_ingest_large` | Excel → SQLite for 1M+ rows |
| **Knowledge** | `excel_search_docs` | Search docs + workbook memory |
| **Knowledge** | `excel_save_query` | Save working code to workbook memory |
| **Corpus** | `excel_get_extraction_prompt` | Get PDF extraction prompt |
| **Corpus** | `excel_extract_to_kb` | Batch-insert extracted entries |
| **Skill** | `excel_read_skill` | Read per-workbook skill documentation |
| **Skill** | `excel_update_skill` | Add/update workbook-specific context and patterns |

### Knowledge Base & Skill Management Tools

#### `excel_search_docs(query, path="", category="", difficulty="", top_k=8)`
Search the knowledge base for documentation and working code examples. Returns:
- Previously working queries for THIS workbook (from query memory)
- Relevant documentation entries ranked by BM25 similarity

Use this when you need to find a pattern that works for your specific workbook type.

#### `excel_save_query(path, description, code, sheet="", columns_used="", result_summary="", tags="")`
Save a successful query to the workbook's memory. This creates a self-improving loop:
1. Search docs → find template → adapt it
2. Analysis succeeds
3. Save working code to workbook memory
4. Next session: this query appears first in `excel_search_docs` results

#### `excel_read_skill(path)`
Read the per-workbook skill document—the LLM's accumulated analytical understanding of a specific Excel file. Returns:
- Context (what the workbook is for)
- Key columns and their quirks
- Validated analytical rules
- Useful analyses that work well
- Warnings about hidden columns, formula issues

#### `excel_update_skill(path, section, content, mode="append", query_key="")`
Update the per-workbook skill document. Use these sections:
- **Context**: Workbook purpose, owner, timeframe (e.g., "FY 2025-26 Nagaland state budget")
- **Key Columns**: Column meanings and special handling rules
- **Analytical Rules**: Domain-specific validation logic
- **Useful Analyses**: Analyses that work well, linked to query keys
- **Warnings**: Hidden columns, formula issues, data quality problems

---

## Knowledge Base Ingestion Pipeline

### Building Your Own Knowledge Base from Documents

The `actual_ingestion.py` script processes PDFs and EPUB books, extracting code examples and analysis patterns into the knowledge base.

```bash
# 1. Place your PDF/EPUB files in the project directory
#    Example: budget_guide.pdf, financial_analysis.epub

# 2. Run the ingestion pipeline
python actual_ingestion.py
```

### What Gets Extracted

- **Python code patterns** from documentation and tutorials
- **Excel formulas** (SUM, VLOOKUP, XLOOKUP, IF, etc.)
- **Financial analysis techniques** (NPV, IRR, PMT calculations)
- **Data manipulation patterns** (groupby, merge, pivot tables)

### Output

The script creates knowledge entries with:
- Automatically detected category (pandas, sql, excel, visualisation, analysis)
- Subcategory (data_manipulation, formula, chart, lookup, etc.)
- Tags for improved searchability
- Source document attribution

---

## Example: Skill Management Workflow

```python
# 1. Initialize a workbook
excel_init("budget.xlsx")

# 2. Read the initial skill documentation
excel_read_skill("budget_ai_workbook.xlsx")
# → Shows empty skill (new workbook)

# 3. Discover patterns during analysis and save them
# After finding a useful chart pattern:
excel_update_skill(
    "budget_ai_workbook.xlsx",
    section="Useful Analyses",
    content="- **Department Utilisation Chart**: Horizontal bar showing % utilisation by dept\n  [query: abc123def]"
)

# 4. Update context for future sessions
excel_update_skill(
    "budget_ai_workbook.xlsx",
    section="Context",
    content="FY 2025-26 Nagaland state budget. 7 departments covering Health, Education, Agriculture."
)

# 5. Next session: read skill to get domain context immediately
excel_read_skill("budget_ai_workbook.xlsx")
# → Now includes your custom context and useful patterns

# 6. Search for relevant patterns in your workbook's memory
excel_search_docs(
    "utilisation chart by department",
    path="budget_ai_workbook.xlsx"
)
# → Returns both global docs AND your saved queries from this workbook
```


## Example Workflow

```python
# 1. Initialize a workbook (creates _ai_workbook.xlsx and analysis artifacts)
excel_init("budget.xlsx")

# 2. Search for relevant patterns in your knowledge base
excel_search_docs("chart revenue by department", path="budget_ai_workbook.xlsx")

# 3. Execute a plain English task
excel_execute(
    "budget_ai_workbook.xlsx",
    "Create a chart showing Expense by Department"
)

# 4. Audit data quality (optional, runs on-demand)
excel_audit("budget_ai_workbook.xlsx")
```

## Configuration

Add to your MCP client configuration:

```json
{
  "mcpServers": {
    "byebyteexcel-v2": {
      "command": "python",
      "args": ["/path/to/ByeByeExcel/server.py"]
    }
  }
}
```

### LLM Skill Placement

- **Claude Code**: `~/.claude/skills/byebyteexcel/SKILL.md`
- **OpenCode**: `~/.opencode/skills/byebyteexcel/SKILL.md`

## File Structure

Per-workbook files are created alongside your `.xlsx`:

| File | Description |
|------|-------------|
| `{name}_ai_workbook.xlsx` | Working copy (original untouched) |
| `{name}_graph.json` | Workbook structure + formatting graph |
| `{name}_skill.md` | Per-workbook skill documentation |
| `{name}_queries.json` | Executable query memory |
| `{name}_analysis/` | Generated charts, reports, exports |

## Dependencies

```
fastmcp           # MCP server framework
openpyxl          # Excel file reading/writing
pandas            # Data manipulation
matplotlib        # Chart generation
rapidfuzz         # Fuzzy string matching
sqlite-utils    # SQLite utilities
```

Optional:
```
oletools          # VBA extraction from .xlsm files
```

## Design Principles

1. **Domain Agnostic**: The MCP exposes primitives; the LLM decides strategy
2. **On-Demand Analysis**: Heavy operations (audit, dedup) run only when requested
3. **Chunked Processing**: Handle 1M+ rows without loading everything into memory
4. **SQLite FTS5 Search**: Zero external dependencies for knowledge search
5. **Transparent Backend**: Automatically switches between pandas and SQLite

## Known Limitations

- Single sheet max: 1,048,576 rows (Excel limitation)
- Formula values show last-computed state (not live recalculated)
- FTS5 stemmer works best with English; Indian languages fall back to character matching
- Full dedup on 1M+ rows takes ~50 minutes (one-time batch process)

## License

MIT License. See [LICENSE](LICENSE) for details.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for upgrade proposals and contribution guidelines.

# Excel MCP — Installation & Setup

## Install

```bash
pip install fastmcp openpyxl pandas matplotlib
pip install oletools   # optional: full VBA extraction from .xlsm
```

## Start the server

```bash
python server.py
```

## Connect to your AI agent

### Claude Code (`~/.claude.json` or `.claude/settings.json`)

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["/absolute/path/to/excel_mcp/server.py"]
    }
  }
}
```

### OpenCode / LM Studio agent

```json
{
  "mcp": {
    "servers": {
      "excel": {
        "command": "python",
        "args": ["/absolute/path/to/excel_mcp/server.py"]
      }
    }
  }
}
```
LM Studio/cline configuration when python server is running is separately. This configuration works best and is tested
```json
"Excel_Master": {
      "url": "http://localhost:6699/mcp",
      "timeout": 1200000
}
```
```json
cline setting
"Excel_Master": {
      "type": "streamableHttp", 
      "url": "http://localhost:6699/mcp",
      "disabled": false,
      "timeout": 1200000
  }
```
### Install the skill

Copy `SKILL.md` to your agent's skills directory:

```bash
# Claude Code
cp SKILL.md ~/.claude/skills/excel-analyst/SKILL.md

# OpenCode
cp SKILL.md ~/.opencode/skills/excel-analyst/SKILL.md
```

## How to use

Just point your agent at any Excel file:

```
analyse /data/Budget_Q2_2026.xlsx
```

```
chart revenue by month in /data/sales.xlsx
```

```
pivot units sold by region in /reports/Q1.xlsx
```

The agent calls `excel_inspect` → understands the structure → calls `excel_execute` → returns results. You never paste data into the chat.

## Three tools

| Tool | What it does |
|---|---|
| `excel_inspect(path)` | Returns full workbook graph (schema, formulas, VBA, dependencies). Always call first. |
| `excel_execute(path, task)` | Runs analysis by task description. Handles 1L+ rows safely. |
| `excel_run_code(path, code)` | Run custom Python. Pre-loaded helpers: `read_sheet`, `read_sheet_chunked`, `save_chart`. |

## Large file handling (1,00,000+ rows)

`excel_execute` automatically uses chunked reads for large files.

For `excel_run_code`, use the pre-loaded helpers:

```python
# Safe for any size
for chunk in read_sheet_chunked("Data", chunksize=50000):
    # process 50,000 rows at a time
    print(chunk["Revenue"].sum())

# Or read just what you need
df = read_sheet("Data", nrows=500)
```

## Output files

Charts and CSVs are written to `<workbook_dir>/excel_analysis/`.


---

## Configuration

Edit `config.py` to set your server port and region:

```python
# config.py

PORT    = 6699      # MCP server port

COUNTRY = "IN"      # "IN" | "US" | "GENERIC"
                    # Controls which name normaliser is suggested by default
                    # for name columns. Does NOT disable other normalisers.
```

**Country / region options:**

| Value | Default name normaliser | Use when |
|---|---|---|
| `"IN"` | `phonetic_names` + `title_names` | South/Southeast Asian names with phonetic spelling variants and honorific titles |
| `"US"` | `us_names` | Western names (US, UK, Canada, Australia) with Mr/Mrs/Dr/Jr titles and Mc/Mac variants |
| `"GENERIC"` | `generic_text` | International or mixed datasets, or when you want to select normalisers manually |

The `COUNTRY` setting only changes which normaliser is **suggested first** — the LLM can always choose any normaliser regardless of this setting.

---

## Normalisers

ByeByeExcel ships with eight normalisers for fuzzy deduplication:

| Normaliser | Best for |
|---|---|
| `phonetic_names` | Names with aspirated consonant variants (Tse/Tshe) and optional clan/community suffixes |
| `title_names` | Names with honorific titles (Dr/Prof/Shri/Smt) that appear inconsistently |
| `us_names` | Western names with Mr/Mrs/Jr/Sr titles and Mc/Mac prefix variants |
| `generic_text` | Place names, category fields, any text column without domain-specific patterns |
| `address` | Address columns — strips unit numbers and generic street words |
| `numeric_range` | Amount or quantity columns with minor rounding differences |
| `date_approx` | Date columns where format or day/month order varies |
| `code_id` | ID and reference number columns with spacing/separator inconsistencies |

---

## Extending Normalisers

ByeByeExcel is designed to be extended for any region or domain.

### Adding rules to an existing normaliser

Open `normalisers.py` and edit the rules list for the normaliser you want to extend.

**Example — adding titles to `title_names` for a French dataset:**

```python
# In normalisers.py, find _TITLE_RULES and add:
_TITLE_RULES = [
    (r'^(mr|mrs|ms|dr|prof|col|brig|lt|capt|sgt|rev|late|'
     r'shri|smt|kumari|sri|srimati|pu|puni|'
     r'm|mme|mlle)\.?\s*', ''),    # ← add French titles here
    ...
]
```

**Example — adding community suffixes to `phonetic_names` for a new region:**

```python
# In normalisers.py, find _PHONETIC_RULES and extend the suffix pattern:
(r'\b(ao|lotha|sumi|angami|...|'
 r'your_suffix_1|your_suffix_2)\b', ''),
```

### Adding a completely new normaliser

1. Create a class that inherits from `BaseNormaliser` in `normalisers.py`
2. Implement `normalise(value)` and optionally override `score(a, b)`
3. Add it to `REGISTRY`
4. Add it to `_COUNTRY_NAME_NORMALISER` or the suggestion logic in `suggest_for_column()`

```python
# normalisers.py

_MY_REGION_RULES = [
    (r'^(title1|title2)\.?\s*', ''),   # strip titles
    (r'suffix_pattern', ''),              # strip suffixes
]

class MyRegionNamesNormaliser(BaseNormaliser):
    name = "my_region_names"
    description = "Name normaliser for [your region]."
    best_for  = ["Name columns in [your region] datasets"]
    not_for   = ["Names without [your region] patterns"]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        s = re.sub(r"[^\w\s]", " ", s)
        for pattern, replacement in _MY_REGION_RULES:
            s = re.sub(pattern, replacement, s, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a": "Title Name Suffix",
            "input_b": "Name",
            "score":   round(self.score("Title Name Suffix", "Name"), 3),
        }

# Add to registry
REGISTRY["my_region_names"] = MyRegionNamesNormaliser()

# Add to suggestion map in config.py:
# COUNTRY = "MY"
# And in normalisers.py _COUNTRY_NAME_NORMALISER:
# "MY": {"primary": "my_region_names", "secondary": "generic_text",
#        "reason": "my_region_names handles ..."}
```

### Using an AI assistant to build a normaliser for your region

Paste the following prompt into any capable LLM (Claude, GPT-4, etc.) to generate a normaliser for your region:

---

```
I am building a fuzzy deduplication normaliser for name data from [YOUR REGION/COUNTRY].

The names in my dataset have these characteristics:
- [Describe title/honorific patterns, e.g. "Names often start with Dr/Mr/Prof"]
- [Describe suffix patterns, e.g. "Clan or community names often appear as a second word"]
- [Describe phonetic variants, e.g. "The same name is spelled with/without doubled letters"]
- [Provide 5-10 example name pairs that should match, e.g. "Müller / Mueller"]

Using this base class from ByeByeExcel's normalisers.py:

class BaseNormaliser:
    def normalise(self, value) -> str: ...   # returns normalised string
    def score(self, a, b) -> float: ...      # returns 0-1 similarity

Please write:
1. A _RULES list of (regex_pattern, replacement) tuples that normalise my names
2. A class MyRegionNamesNormaliser(BaseNormaliser) implementing normalise()
3. An _example() dict showing two name variants and their score
4. Instructions for adding it to REGISTRY and _COUNTRY_NAME_NORMALISER

The normaliser should handle the most common data entry inconsistencies
for this region while avoiding false positives (matching different people).
```

---

Replace `[YOUR REGION/COUNTRY]` and the bullet points with your actual data characteristics. The more specific your examples, the better the generated normaliser will be.