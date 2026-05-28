"""
ingest.py — Universal column mapping + large-file SQLite ingestion.

Two public functions:
  map_columns(df_columns)           → {original_col: canonical_field}
  ingest_excel_to_sqlite(path, db)  → ingestion report dict

Design:
  - One SQLite table per sheet (default). Each table named from the
    sanitised sheet name: "Revenue Data" → "revenue_data".
  - merge_sheets=True for same-schema split files (e.g. 11L register
    split across Sheet1/Sheet2/Sheet3). All sheets → one table.
  - Freshness check on re-ingest: if Excel mtime <= last ingested_at,
    skip and return "DB is current". Pass overwrite=True to force rebuild.
  - _sheet_registry table maps sheet_name → table details for StorageBackend.
  - Master schema pre-scan (change 3, already applied) prevents schema
    mismatch crashes on merge_sheets=True mode only.
"""

import re, hashlib, unicodedata, sqlite3, json
from pathlib import Path
from datetime import datetime
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# CANONICAL FIELD PATTERNS
# ─────────────────────────────────────────────────────────────────────────────

CANONICAL_MAP = {
    "name": [
        r"^name$", r"beneficiar\w*\s*name", r"name\s*of\s*beneficiar\w*",
        r"applicant\s*name", r"member\s*name", r"household\s*head",
        r"full\s*name", r"recipient\s*name", r"claimant", r"individual\s*name",
        r"account\s*holder\s*name", r"person\s*name",
    ],
    "father_name": [
        r"father", r"\bf/o\b", r"\bs/o\b", r"son\s*of",
        r"father'?s?\s*name", r"guardian\s*name", r"parent\s*name",
    ],
    "husband_name": [r"husband", r"\bh/o\b", r"\bw/o\b", r"spouse"],
    "village":      [r"village", r"\bgram\b", r"\bgaon\b", r"habitation",
                     r"locality", r"ward\s*name", r"hamlet"],
    "district":     [r"district", r"\bdist\b", r"\bzila\b", r"\bjilla\b"],
    "block":        [r"\bblock\b", r"mandal", r"tehsil", r"taluk", r"\bpanchayat\b"],
    "state":        [r"^state$", r"state\s*name"],
    "dob":          [r"date\s*of\s*birth", r"\bdob\b", r"birth\s*date", r"d\.o\.b",
                     r"janm\s*tithi"],
    "age":          [r"^age$", r"\bage\b", r"years?\s*old", r"\bumar\b"],
    "gender":       [r"gender", r"\bsex\b", r"m/f", r"male.female"],
    "category":     [r"category", r"caste", r"\btribe\b", r"community", r"sc/st"],
    "aadhaar":      [r"aadhaar", r"aadhar", r"\buid\b", r"\buidai\b", r"adhaar",
                     r"aadhar\s*no", r"uid\s*no", r"aadhaar\s*number"],
    "mobile":       [r"mobile", r"phone", r"contact\s*no", r"\bcell\b",
                     r"\bmob\b", r"mo\.?\s*no", r"telephone"],
    "bank_account": [r"account\s*no", r"bank\s*acc", r"a/c\s*no", r"\bacct\b",
                     r"account\s*number"],
    "ifsc":         [r"\bifsc\b", r"ifsc\s*code"],
    "bank_name":    [r"bank\s*name", r"name\s*of\s*bank", r"\bbank\b(?!\s*acc)"],
    "amount":       [r"amount", r"sanctioned", r"payment", r"benefit",
                     r"installment", r"disbursed", r"released", r"grant"],
    "status":       [r"^status$", r"payment\s*status", r"verification\s*status",
                     r"approval\s*status"],
    "beneficiary_id": [r"beneficiar\w*\s*id", r"application\s*no", r"reg\w*\s*no",
                       r"registration\s*id", r"^sl\.?\s*no", r"^s\.?\s*no",
                       r"ref\s*no", r"case\s*no"],
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _norm_header(h: str) -> str:
    if not isinstance(h, str): return ""
    h = unicodedata.normalize("NFKD", h)
    h = "".join(c for c in h if not unicodedata.combining(c))
    return re.sub(r"[^\w\s/]", " ", h.lower()).strip()


def map_columns(df_columns: list) -> dict:
    """Map actual column names → canonical field names. First match wins."""
    mapping: dict = {}
    used: set = set()
    for col in df_columns:
        norm = _norm_header(str(col))
        for canonical, patterns in CANONICAL_MAP.items():
            if canonical not in used:
                if any(re.search(p, norm) for p in patterns):
                    mapping[col] = canonical
                    used.add(canonical)
                    break
    return mapping


def _clean(val):
    if val is None: return None
    s = str(val).strip()
    return None if s in ("", "nan", "NaN", "NULL", "null", "None",
                         "-", "N/A", "NA", "#N/A", "NIL", "nil") else s


def _sanitise_table_name(sheet_name: str) -> str:
    """Convert sheet name to a valid SQLite table name."""
    s = str(sheet_name).lower().strip()
    s = re.sub(r"[^\w]", "_", s)          # non-word chars → underscore
    s = re.sub(r"_+", "_", s).strip("_")  # collapse multiple underscores
    if not s or s[0].isdigit():
        s = "sheet_" + s                   # can't start with digit
    return s or "sheet"


def _infer_sqlite_type(series: pd.Series) -> str:
    """Infer SQLite column type from a pandas Series sample."""
    non_null = series.dropna()
    if non_null.empty: return "TEXT"
    try:
        pd.to_numeric(non_null)
        if (non_null.astype(str).str.contains(r"\.")).any():
            return "REAL"
        return "INTEGER"
    except Exception:
        return "TEXT"


def _compute_join_candidates(sheet_cols: dict) -> dict:
    """
    For each sheet, identify columns that appear in 2+ sheets —
    these are natural join keys.
    sheet_cols: {sheet_name: [col_name, ...]}
    Returns: {sheet_name: [join_candidate_col, ...]}
    """
    from collections import Counter
    all_cols = []
    for cols in sheet_cols.values():
        all_cols.extend(cols)
    freq = Counter(all_cols)
    shared = {c for c, n in freq.items() if n >= 2}

    return {
        sheet: [c for c in cols if c in shared]
        for sheet, cols in sheet_cols.items()
    }


# ─────────────────────────────────────────────────────────────────────────────
# FRESHNESS CHECK
# ─────────────────────────────────────────────────────────────────────────────

def _is_db_current(excel_path: str, db_path: str) -> bool:
    """
    Returns True if the db exists and was ingested AFTER the Excel was
    last modified — meaning the db is up to date and can be reused.
    """
    if not Path(db_path).exists():
        return False
    try:
        excel_mtime = Path(excel_path).stat().st_mtime
        conn = sqlite3.connect(db_path)
        row = conn.execute(
            "SELECT ingested_at FROM _ingest_log ORDER BY ingested_at DESC LIMIT 1"
        ).fetchone()
        conn.close()
        if not row:
            return False
        ingested_at = datetime.fromisoformat(row[0]).timestamp()
        return excel_mtime <= ingested_at
    except Exception:
        return False


def _drop_all_tables(conn: sqlite3.Connection):
    """Drop all user tables from an existing db for a full rebuild."""
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
    for (tname,) in cur.fetchall():
        cur.execute(f"DROP TABLE IF EXISTS [{tname}]")
    conn.commit()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN INGESTION
# ─────────────────────────────────────────────────────────────────────────────

def ingest_excel_to_sqlite(
    excel_path:   str,
    db_path:      str,
    chunk_size:   int  = 10_000,
    overwrite:    bool = False,
    merge_sheets: bool = False,
) -> dict:
    """
    Stream an Excel workbook into SQLite.

    Default (merge_sheets=False):
      One table per sheet, named from the sanitised sheet name.
      "Expenditure" → table expenditure
      "Revenue Data" → table revenue_data
      Use this for workbooks where sheets have different schemas
      (budget workbooks, multi-entity files).

    merge_sheets=True:
      All sheets combined into one table named 'data'.
      Use this for same-schema files split across sheets
      (11L beneficiary register split as Sheet1/Sheet2/Sheet3).

    Freshness check:
      If db exists and Excel mtime <= last ingested_at → returns
      {"status": "current", ...} without re-ingesting.
      Pass overwrite=True to force a full rebuild.

    Returns ingestion report dict including sheet_tables mapping.
    """
    p = Path(excel_path)

    # ── Freshness check ───────────────────────────────────────────────────────
    if not overwrite and _is_db_current(excel_path, db_path):
        # Read existing registry to return full report
        try:
            conn = sqlite3.connect(db_path)
            rows = conn.execute(
                "SELECT sheet_name, table_name, row_count, columns_json "
                "FROM _sheet_registry"
            ).fetchall()
            conn.close()
            sheet_tables = {
                r[0]: {
                    "table":          r[1],
                    "rows":           r[2],
                    "columns":        json.loads(r[3]) if r[3] else [],
                    "canonical_cols": {},
                    "join_candidates": [],
                }
                for r in rows
            }
            return {
                "status":       "current",
                "message":      (f"DB is current. Excel not modified since last "
                                 f"ingest. Pass overwrite=True to force rebuild."),
                "db_path":      db_path,
                "sheet_tables": sheet_tables,
                "total_rows":   sum(r[2] for r in rows),
            }
        except Exception:
            pass  # fall through to re-ingest

    # ── Full rebuild ──────────────────────────────────────────────────────────
    conn = sqlite3.connect(db_path)
    if overwrite or Path(db_path).exists():
        _drop_all_tables(conn)

    xl         = pd.ExcelFile(excel_path)
    total_rows = 0
    all_maps:  dict = {}
    sheet_stats: dict = {}

    # Collect headers per sheet for join candidate computation
    sheet_raw_cols: dict = {}
    for sheet in xl.sheet_names:
        hdf = pd.read_excel(excel_path, sheet_name=sheet, nrows=0)
        sheet_raw_cols[sheet] = list(hdf.columns)

    join_candidates_map = _compute_join_candidates(sheet_raw_cols)

    # ── merge_sheets=True: master schema pre-scan (existing logic) ────────────
    master_column_list = None
    if merge_sheets:
        master_cols: set = set()
        for sheet in xl.sheet_names:
            orig_cols  = sheet_raw_cols[sheet]
            col_map_pre = map_columns(orig_cols)
            master_cols.update(orig_cols)
            master_cols.update([f"_c_{c}" for c in col_map_pre.values()])
        master_cols.update(["_source_sheet", "_source_row", "_row_hash"])
        master_column_list = list(master_cols)

    # ── Sheet-level registry data (built during ingestion) ────────────────────
    registry_rows: list = []

    for sheet in xl.sheet_names:
        orig_cols = sheet_raw_cols[sheet]
        col_map   = map_columns(orig_cols)
        all_maps[sheet] = col_map

        table_name   = "data" if merge_sheets else _sanitise_table_name(sheet)
        table_exists = False
        headers      = orig_cols
        skip         = 1
        sheet_rows   = 0

        # Sample first data row for type inference
        sample_df = pd.read_excel(excel_path, sheet_name=sheet,
                                  nrows=5, header=0)

        # Build column details for registry
        col_details = []
        for col in orig_cols:
            sqlite_type = _infer_sqlite_type(
                sample_df[col] if col in sample_df.columns else pd.Series(dtype=object))
            col_details.append({"name": col, "type": sqlite_type})
        # Add canonical cols
        for orig, canon in col_map.items():
            col_details.append({"name": f"_c_{canon}", "type": "TEXT"})
        # Add metadata cols
        for mc in ["_source_sheet", "_source_row", "_row_hash"]:
            col_details.append({"name": mc, "type": "TEXT"})

        # ── Chunked ingestion ─────────────────────────────────────────────────
        while True:
            try:
                chunk = pd.read_excel(
                    excel_path, sheet_name=sheet,
                    header=None, names=headers,
                    skiprows=skip, nrows=chunk_size,
                )
            except Exception:
                break

            if chunk.empty:
                break

            chunk["_source_sheet"] = sheet
            chunk["_source_row"]   = range(skip + 1, skip + 1 + len(chunk))

            # Canonical columns
            for orig, canon in col_map.items():
                if orig in chunk.columns:
                    chunk[f"_c_{canon}"] = chunk[orig].apply(_clean)

            # Row hash
            c_cols = [f"_c_{c}" for c in col_map.values()
                      if f"_c_{c}" in chunk.columns]
            if c_cols:
                chunk["_row_hash"] = chunk[c_cols].apply(
                    lambda r: hashlib.md5(
                        "|".join(str(v or "") for v in r).encode()
                    ).hexdigest(), axis=1)

            # Clean non-meta columns
            for col in chunk.columns:
                if not col.startswith("_"):
                    chunk[col] = chunk[col].apply(_clean)

            # merge_sheets: align to master schema
            if merge_sheets and master_column_list:
                chunk = chunk.reindex(columns=master_column_list)

            safe_db_chunksize = max(1, 990 // len(chunk.columns))
            chunk.to_sql(
                table_name, conn,
                if_exists="append" if table_exists else "replace",
                index=False, method="multi",
                chunksize=safe_db_chunksize,
            )
            table_exists = True
            sheet_rows  += len(chunk)
            total_rows  += len(chunk)
            skip        += chunk_size

            if len(chunk) < chunk_size:
                break

        sheet_stats[sheet] = sheet_rows

        # Indexes on canonical columns for this table
        cur = conn.cursor()
        for field in CANONICAL_MAP:
            col = f"_c_{field}"
            try:
                cur.execute(
                    f"CREATE INDEX IF NOT EXISTS "
                    f"idx_{table_name}_{field} ON [{table_name}]({col})"
                )
            except Exception:
                pass
        try:
            cur.execute(
                f"CREATE INDEX IF NOT EXISTS "
                f"idx_{table_name}_hash ON [{table_name}](_row_hash)"
            )
        except Exception:
            pass

        registry_rows.append((
            sheet,
            table_name,
            sheet_rows,
            json.dumps(col_details),
            json.dumps([c for c in (join_candidates_map.get(sheet) or [])]),
            json.dumps({canon: orig
                        for orig, canon in col_map.items()}),
        ))

    # ── Metadata tables ───────────────────────────────────────────────────────
    cur = conn.cursor()

    # Sheet registry — the lookup table StorageBackend uses
    cur.execute("""CREATE TABLE IF NOT EXISTS _sheet_registry (
        sheet_name      TEXT PRIMARY KEY,
        table_name      TEXT NOT NULL,
        row_count       INTEGER,
        columns_json    TEXT,
        join_candidates TEXT,
        canonical_map   TEXT
    )""")
    cur.executemany(
        "INSERT OR REPLACE INTO _sheet_registry VALUES (?,?,?,?,?,?)",
        registry_rows
    )

    # Column map (existing — kept for compatibility)
    cur.execute("""CREATE TABLE IF NOT EXISTS _column_map
        (sheet TEXT, original_col TEXT, canonical_col TEXT)""")
    for sheet, cmap in all_maps.items():
        for orig, canon in cmap.items():
            cur.execute("INSERT INTO _column_map VALUES(?,?,?)",
                        (sheet, orig, canon))

    # Ingest log
    cur.execute("""CREATE TABLE IF NOT EXISTS _ingest_log (
        source_file TEXT,
        ingested_at TEXT,
        total_rows  INTEGER,
        sheets      TEXT,
        merge_sheets INTEGER
    )""")
    cur.execute(
        "INSERT INTO _ingest_log VALUES(?,?,?,?,?)",
        (str(p), datetime.utcnow().isoformat(),
         total_rows, json.dumps(xl.sheet_names), int(merge_sheets))
    )

    conn.commit()
    conn.close()

    # ── Build sheet_tables return value (rich — for graph + tool output) ──────
    sheet_tables = {}
    for (sheet, table_name, row_count,
         cols_json, jc_json, canon_json) in registry_rows:
        sheet_tables[sheet] = {
            "table":           table_name,
            "rows":            row_count,
            "columns":         json.loads(cols_json),
            "join_candidates": json.loads(jc_json),
            "canonical_cols":  json.loads(canon_json),
        }

    return {
        "status":       "ingested",
        "total_rows":   total_rows,
        "sheets":       sheet_stats,
        "column_maps":  all_maps,
        "db_path":      db_path,
        "merge_sheets": merge_sheets,
        "sheet_tables": sheet_tables,
    }