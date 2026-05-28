## Purpose

This document is written for a coding agent (Claude Code, OpenCode, or similar).
Your task is to read authoritative textbooks and reference materials on Excel,
pandas, data analysis, and SQL, then populate the ByeByeExcel knowledge base
(`excel_knowledge.db`) with working, tested code examples.

The knowledge base powers `excel_search_docs` — when an LLM doesn't know how
to do something, it searches this corpus and gets back ranked, working code it
can immediately adapt. The richer the corpus, the more capable the LLM becomes
without hallucinating.

---

## What the corpus needs to contain

Each entry answers a specific, real question a data analyst would ask.
Not documentation summaries. Not concept explanations. Working code.

**Good entry title:** "Rolling 12-month average grouped by region"
**Bad entry title:** "pandas window functions overview"

The entry must have:
- A concrete problem statement (plain English)
- Code that runs without modification (except swapping column names)
- Notes about gotchas or edge cases
- Tags that match how someone would search for it

---

## Entry format (JSON, one per entry)

```json
{
  "id":          "unique-slug-001",
  "category":    "pandas",
  "subcategory": "groupby",
  "title":       "Rolling 12-month sum by category",
  "problem":     "Compute a moving 12-month total for each category in a monthly dataset",
  "tags":        ["rolling", "window", "12-month", "monthly", "category"],
  "code":        "df = read_sheet(sheet_name)\ndf['rolling_12m'] = df.groupby('Category_Column')['Amount_Column'].transform(lambda x: x.rolling(12, min_periods=1).sum())\nprint(df.head(20).to_string())",
  "notes":       "transform() broadcasts the result back to original DataFrame shape. Use min_periods=1 to avoid NaN at the start of each group.",
  "source":      "Python for Data Analysis — Wes McKinney, 3rd Ed, Ch.11",
  "difficulty":  "intermediate"
}
```

**Category values:** `pandas` | `sql` | `analysis` | `visualisation` | `excel` | `dedup`

**Subcategory values by category:**
- pandas:        `groupby` | `pivot` | `filter` | `window` | `merge` | `date` | `string` | `cleaning` | `formatting`
- sql:           `groupby` | `join` | `window` | `anomaly` | `filter` | `aggregation`
- analysis:      `financial` | `distribution` | `statistical` | `time-series` | `ranking`
- visualisation: `bar` | `line` | `scatter` | `heatmap` | `chart` | `dashboard`
- excel:         `formula` | `formatting` | `pivot` | `vba` | `named-range`
- dedup:         `names` | `ids` | `strategy` | `blocking` | `scoring`

**Difficulty:** `beginner` | `intermediate` | `advanced`

---

## Priority sources to process

### Tier 1 — Highest priority

**1. Python for Data Analysis — Wes McKinney (3rd Edition)**
Chapters: 5 (pandas basics), 8 (join/merge), 10 (groupby), 11 (time series)
Categories: `pandas/groupby`, `pandas/merge`, `pandas/window`, `pandas/date`
Target: 200+ entries

**2. pandas Documentation (pandas.pydata.org)**
Sections: GroupBy, Reshaping, Time Series, IO tools
Categories: `pandas/pivot`, `pandas/filter`, `pandas/cleaning`, `pandas/formatting`
Target: 150+ entries

**3. Python Data Science Handbook — Jake VanderPlas**
Chapters: Ch.3 (data manipulation), Ch.4 (visualisation)
Categories: `pandas/filter`, `visualisation/chart`, `analysis/statistical`
Target: 100+ entries

**4. SQL for Data Analysis — Cathy Tanimura (O'Reilly)**
Focus: aggregations, window functions, self-joins, cohort analysis
Categories: `sql/groupby`, `sql/window`, `sql/aggregation`
Target: 80+ entries

### Tier 2 — Important domain knowledge

**5. Microsoft Excel Bible — John Walkenbach**
Focus: formula patterns, array formulas, pivot tables, named ranges
Categories: `excel/formula`, `excel/pivot`, `excel/formatting`
Target: 100+ entries

**6. Storytelling with Data — Cole Nussbaumer Knaflic**
Focus: chart type selection, clarity patterns, annotation
Categories: `visualisation/bar`, `visualisation/line`, `visualisation/chart`
Target: 50+ entries

**7. Practical Statistics for Data Scientists**
Focus: distributions, outlier detection, correlation, hypothesis testing
Categories: `analysis/statistical`, `analysis/distribution`
Target: 60+ entries

### Tier 3 — Specialised

**8. Financial analysis patterns (various sources)**
Focus: budget vs actual, utilisation rates, variance analysis, period comparisons
Categories: `analysis/financial`
Target: 40+ entries

**9. Data quality patterns (various sources)**
Focus: deduplication strategies, anomaly detection, data cleaning
Categories: `dedup/strategy`, `pandas/cleaning`
Target: 50+ entries

---

## How to add entries

### Method 1: Single entry via MCP tool
```python
excel_add_doc_entry(
    title       = "Your title",
    problem     = "What problem this solves",
    code        = "working python code here",
    category    = "pandas",
    subcategory = "groupby",
    tags        = "tag1,tag2,tag3",
    notes       = "gotcha or edge case",
    source      = "Book Name, Chapter X",
    difficulty  = "intermediate",
)
```

### Method 2: Bulk JSON insert
```python
import sys
sys.path.insert(0, '/path/to/excel_mcp')
from knowledge import init_knowledge_db, bulk_insert

db_path = init_knowledge_db()
entries = [
    {
        "id": "your-id-001",
        "category": "pandas", "subcategory": "groupby",
        "title": "...", "problem": "...",
        "tags": ["tag1", "tag2"],
        "code": "...", "notes": "...",
        "source": "Book, Ch.X", "difficulty": "intermediate",
    },
]
result = bulk_insert(entries, db_path)
print(f"Inserted: {result['inserted']}, Skipped: {result['skipped']}")
```

### Method 3: From a structured markdown file
```python
import re, sys
sys.path.insert(0, '/path/to/excel_mcp')
from knowledge import bulk_insert

def parse_example_md(md_text):
    """
    Each entry in the markdown file:
    ## Title
    **Problem:** description
    **Category:** pandas/groupby
    **Tags:** tag1, tag2, tag3
    **Difficulty:** intermediate
    **Source:** Book, Ch.X
    ```python
    code here
    ```
    **Notes:** gotcha
    ---
    """
    entries = []
    for block in md_text.split("\n---\n"):
        if not block.strip(): continue
        try:
            title    = re.search(r"^## (.+)", block, re.M).group(1).strip()
            problem  = re.search(r"\*\*Problem:\*\* (.+)", block).group(1).strip()
            cat_str  = re.search(r"\*\*Category:\*\* (.+)", block).group(1).strip()
            tags_str = re.search(r"\*\*Tags:\*\* (.+)", block).group(1).strip()
            diff     = re.search(r"\*\*Difficulty:\*\* (.+)", block).group(1).strip()
            source   = re.search(r"\*\*Source:\*\* (.+)", block).group(1).strip()
            code     = re.search(r"```python\n(.+?)```", block, re.DOTALL).group(1).strip()
            notes_m  = re.search(r"\*\*Notes:\*\* (.+)", block)
            notes    = notes_m.group(1).strip() if notes_m else ""
            cat_parts   = cat_str.split("/")
            entries.append({
                "category":    cat_parts[0],
                "subcategory": cat_parts[1] if len(cat_parts) > 1 else "",
                "title": title, "problem": problem,
                "tags":  [t.strip() for t in tags_str.split(",")],
                "code":  code, "notes": notes,
                "source": source, "difficulty": diff,
            })
        except Exception as e:
            print(f"Skipped block: {e}")
    return entries

md_text = open("your_extracted_entries.md").read()
result  = bulk_insert(parse_example_md(md_text))
print(f"Inserted {result['inserted']} entries")
```

---

## What makes a good entry

### Code quality rules

1. **Use the pre-loaded helpers** — `read_sheet()`, `read_all_sheets()`, `save_chart()`,
   `save_new_workbook()`, `AI_WORKBOOK`, `OUTPUT_DIR`, `GRAPH`
2. **Use placeholder names** — `"Category_Column"`, `"Numeric_Col"` so LLM substitutes from graph
3. **Every code block must be complete and runnable** — the LLM pastes it into `excel_run_code`
4. **Handle the percentage column gotcha** — if code touches a `%` column, add the ×100 note
5. **Include a print or save_chart call** — LLM needs visible output to confirm it worked

### Negative examples (do not add)

```
# BAD: concept explanation without code
"pivot_table() reshapes data by moving unique values into columns"

# BAD: imports not available in the sandbox
import seaborn as sns
sns.heatmap(...)

# BAD: hardcoded column names
df['Budget_Approved_2024'].sum()

# GOOD: placeholder the LLM replaces from the graph
col = "Amount_Column"   # replace with actual column from graph
df[col].sum()
```

---

## Quality targets

| Category | Target entries | Priority |
|---|---|---|
| pandas/groupby | 80 | Critical |
| pandas/window | 40 | Critical |
| pandas/filter | 50 | High |
| pandas/date | 40 | High |
| pandas/merge | 30 | High |
| pandas/cleaning | 40 | High |
| analysis/financial | 50 | High |
| visualisation/* | 60 | Medium |
| sql/* | 60 | Medium |
| excel/* | 50 | Medium |
| dedup/* | 30 | Medium |
| **TOTAL** | **530+** | |

The seed corpus (`seed_corpus.py`) provides ~30 entries as a starting point.

---

## Verification

After bulk insert, verify quality:
```python
from knowledge import search_docs

test_queries = [
    "rolling average by month",
    "budget vs actual comparison",
    "find outliers above average",
    "merge two sheets on common column",
    "percentage column display correctly",
    "group by region sum amount",
    "waterfall chart variance",
    "phonetic name deduplication",
    "large file sql group by",
]
for q in test_queries:
    results = search_docs(q, top_k=3)
    top    = results[0] if results else None
    status = "✅" if top else "❌ NO RESULTS"
    title  = top['title'] if top else "—"
    print(f"{status}  '{q}' -> {title}")
```

All test queries should return at least one relevant result.