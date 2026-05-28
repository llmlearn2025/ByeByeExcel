# ByeByeExcel
Solving AI Limitations on Large Excel Analysis (>50k Rows): A Metadata-Driven Architecture.
# ByeByeExcel

An intelligent Excel analysis and data processing system that brings LLM-powered insights directly from your spreadsheets.

## Overview

ByeByeExcel is a Model Context Protocol (MCP) server that gives LLM agents (Claude Code, OpenCode, Cursor, LM Studio) the ability to work intelligently with Excel files—including massive datasets with 1+ million rows.

**Core Philosophy:** The MCP provides execution primitives and file management, while the LLM handles intelligent decision-making, strategy selection, and result interpretation.

## Features

- **Intelligent Workbook Analysis** - Automatic structure scanning, format detection, and column type classification
- **Plain English Commands** - Execute analysis tasks using natural language ("summarise Revenue by Region")
- **Large File Support** - Handle 1M+ row files with SQLite-backed chunked processing
- **Data Quality Auditing** - Detect formula overrides, logic violations, type errors, and consistency issues
- **Fuzzy Deduplication** - Smart duplicate detection with configurable normalisers for names, dates, codes, and text
- **Knowledge Management** - Per-workbook query memory and shared knowledge base for reusable patterns
- **Rich Output** - Automatic output tiering (Markdown tables, HTML previews, charts, exports)

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