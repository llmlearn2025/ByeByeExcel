## Core principle

**The LLM is the intelligence. The MCP is the execution engine.**

The MCP never decides what columns to match on, which normaliser fits the data,
or what thresholds make sense. It provides tools and primitives. The LLM decides
how to use them after reading the graph.

---

## Always start here

### Step 1 — `excel_init(path)`

Call once per file. Returns the complete workbook graph.

Read the graph carefully before anything else. It tells you:
- Exact sheet names and headers (use verbatim)
- Column types: `numeric`, `text`, `date`, `boolean`
- Column formats: `percentage`, `currency`, `date`, `number_thousands`
- Fill meanings: `input_hardcoded` (blue), `link_cross_sheet` (green),
  `assumption` (yellow), `output_display` (grey)
- Hidden columns (⚠️ HIDDEN → exclude from aggregations)
- Formula patterns and cross-sheet dependencies
- VBA module source (for .xlsm)

---

## Analysis tools

| User wants | Tool |
|---|---|
| Summary / statistics | `excel_execute(ai_path, "summarise")` |
| Chart | `excel_execute(ai_path, "chart Revenue by Month")` |
| Pivot | `excel_execute(ai_path, "pivot Units by Region")` |
| Totals | `excel_execute(ai_path, "total all numeric columns")` |
| All sheets | `excel_execute(ai_path, "all sheets overview")` |
| Custom logic | `excel_run_code(ai_path, code)` |

---

## Fuzzy dedup — the intelligence split

### Step 1 — ask what normalisers exist
```
excel_get_fuzzy_options(ai_path, "Name_Column,Secondary_Column")
```
Returns per-column suggestions based on graph data (format, fill, sample values).
Read these. Decide which normaliser fits each column.

### Step 2 — decide (LLM's job, not the MCP's)

Ask yourself:
- Are names written with aspirated consonant variants or clan/tribal suffixes? → `phonetic_names`
- Are names written with honorific titles (Dr/Prof/Mr/Mrs)? → `title_names`
- Western-style names (US/UK/AU)? → `us_names`
- Place or location names? → `generic_text`
- Amount / numeric column? → `numeric_range`
- Date of birth? → `date_approx`
- ID / reference number? → `code_id`

The default normaliser for name columns is set by `COUNTRY` in `config.py`.
You can always override — the suggestion is a starting point, not a constraint.

### Step 3 — run dedup with your chosen normalisers
```python
excel_fuzzy_dedup(
    path="...ai_workbook.xlsx",
    columns="Name_Column,Secondary_Column,Location_Column",
    normalisers="phonetic_names,phonetic_names,generic_text",
    weights="0.6,0.25,0.15",
    high_threshold=0.80,
    review_threshold=0.60,
    possible_threshold=0.40,
)
```

The MCP executes. It never second-guesses your normaliser choice.

### For large files with location columns — use run_code instead

When your data has location columns (city, region, district) alongside name
columns, `excel_run_code` with geographic blocking is faster than
`excel_fuzzy_dedup` for files over 10K rows:

```python
# Fast fuzzy dedup: block by location + name prefix
import pandas as pd
from difflib import SequenceMatcher

df = read_sheet(sheet_name)
df["_block"] = (
    df["Location_Column"].astype(str) + "_" +
    df["Name_Column"].astype(str).str[:3].str.upper()
)
matches = []
for _, group in df.groupby("_block"):
    if len(group) < 2: continue
    records = group.to_dict("records")
    for i in range(len(records)):
        for j in range(i+1, len(records)):
            r1, r2 = records[i], records[j]
            ns = SequenceMatcher(None, str(r1["Name_Column"]),
                                       str(r2["Name_Column"])).ratio()
            if ns < 0.80: continue
            ps = SequenceMatcher(None, str(r1["Secondary_Column"]),
                                       str(r2["Secondary_Column"])).ratio()
            if ps >= 0.70:
                matches.append({"Name_A": r1["Name_Column"],
                                 "Name_B": r2["Name_Column"],
                                 "Name_Sim": round(ns,3),
                                 "Parent_Sim": round(ps,3)})
if matches:
    result = pd.DataFrame(matches).sort_values("Name_Sim", ascending=False)
    print(f"Found {len(result):,} pairs")
    path = save_new_workbook(result, "Fuzzy_Matches")
    print("Saved:", path)
```

### Threshold guidance
- Start with defaults (0.80 / 0.60 / 0.40)
- Too many false positives → raise high_threshold to 0.85
- Strict/regulated data → raise possible_threshold to 0.50
- Loose/inconsistent data → lower review_threshold to 0.55
- `block_chars=4` for tighter blocking on large files

### Dedup works on ANY columns
Not just names. Examples:
- Vendor dedup: `columns="Company_Name,Tax_ID"`, `normalisers="generic_text,code_id"`
- Product dedup: `columns="SKU,Description"`, `normalisers="code_id,generic_text"`
- Transaction dedup: `columns="Amount,Date,Account"`, `normalisers="numeric_range,date_approx,code_id"`

---

## Data quality tools (on demand only)

### `excel_audit` — NEVER runs automatically
Call only when user explicitly asks: "audit", "check data quality", "find errors"

### `excel_ingest_large` — for large files (>50K rows)
Converts Excel to SQLite once. Creates one table per sheet.
After that, SQL queries and dedup run against the database.
Pass `merge_sheets=True` if all sheets have the same schema.

---

## Large file rules (> 50K rows)

In `excel_run_code`, always use chunked reads:
```python
for chunk in read_sheet_chunked("Sheet1", chunksize=50000):
    total += chunk["Revenue"].fillna(0).sum()
```

`excel_execute` handles this automatically based on row count from the graph.

---

## Cross-sheet SQL queries (after ingest)

Use `GRAPH["sheet_tables"]` to get table names and join keys:

```python
import sqlite3, pandas as pd
conn = sqlite3.connect(GRAPH["db_path"])
# GRAPH["sheet_tables"] has table name, columns, and join_candidates per sheet
t1 = GRAPH["sheet_tables"]["Sheet1"]["table"]
t2 = GRAPH["sheet_tables"]["Sheet2"]["table"]
jk = GRAPH["sheet_tables"]["Sheet1"]["join_candidates"][0]
df = pd.read_sql(f"""
    SELECT a.*, b.Extra_Column
    FROM [{t1}] a
    LEFT JOIN [{t2}] b ON a.[{jk}] = b.[{jk}]
""", conn)
conn.close()
print(df.head().to_string())
```

---

## Format interpretation rules

- `format=percentage` → values are 0–1 decimals. Multiply ×100 to display.
- `format=date` → may be stored as integer serial. Use `pd.to_datetime` if needed.
- `fill=input_hardcoded` → safe to modify (blue cells)
- `fill=link_cross_sheet` → formula-driven, modify the formula not the value
- `fill=assumption` → flag to user before changing (yellow cells)
- ⚠️ HIDDEN → skip in aggregations unless user explicitly asks for it

---

## Output locations

All charts, new workbooks, and exports go to `<stem>_analysis/`.
Always tell the user the full path of every file produced.

---

## What NOT to do

- Never touch the original file after `excel_init`
- Never guess column names or sheet names — use the graph
- Never hardcode a normaliser — always call `excel_get_fuzzy_options` first
- Never call `excel_audit` without user asking
- Never load all rows into memory for large files