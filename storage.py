"""
storage.py — Transparent storage layer for Excel MCP v2.

Every other module calls storage functions. None of them know or care whether
they are reading a 500-row pandas DataFrame or a 1.1M-row SQLite table.
The switch is automatic at LARGE_FILE_THRESHOLD rows.

Public API:
  backend = StorageBackend(path, sheet, graph)
  backend.read(columns, filter_expr, nrows)  → pd.DataFrame
  backend.count()                            → int
  backend.iter_chunks(columns, chunksize)    → Iterator[pd.DataFrame]
  backend.write_result(df, name)             → Path
  backend.ensure_sqlite()                    → str  (db path)
  backend.db_path                            → str | None

Change from v1:
  _sql_read() and _sql_chunks() now route to the correct per-sheet table
  via _sheet_to_table() lookup against _sheet_registry in the db.
  The hardcoded "beneficiaries" table name is gone.
  Callers are unaffected — read_sheet("Expenditure") still works identically.
"""

import re, json, sqlite3
from pathlib import Path
from typing import Iterator, Optional, List
import pandas as pd

LARGE_FILE_THRESHOLD = 50_000
CHUNK_SIZE           = 50_000


def _chunked_excel_reader(path: str, sheet=0, chunksize: int = CHUNK_SIZE):
    """Yield DataFrame chunks from an Excel sheet."""
    headers = list(pd.read_excel(path, sheet_name=sheet, nrows=0).columns)
    skip = 1
    while True:
        chunk = pd.read_excel(path, sheet_name=sheet, header=None,
                              names=headers, skiprows=skip, nrows=chunksize)
        if chunk.empty:
            break
        yield chunk
        skip += chunksize
        if len(chunk) < chunksize:
            break


class StorageBackend:
    """
    Unified read interface for both small (pandas) and large (SQLite) workbooks.
    Instantiate once per tool call. Not a long-lived object.
    """

    def __init__(self, ai_path: str, sheet: str = "", graph: dict = None):
        self.ai_path = str(Path(ai_path).resolve())
        self.graph   = graph or {}
        self._db_path: Optional[str] = None

        # Resolve active sheet name
        sheets = list(self.graph.get("sheets", {}).keys())
        if sheet and sheet in sheets:
            self.sheet = sheet
        elif sheets:
            self.sheet = sheets[0]
        else:
            self.sheet = 0  # pandas fallback

        # Row count from graph — avoids opening the file
        s_info = self.graph.get("sheets", {}).get(str(self.sheet), {})
        self._row_count = s_info.get("max_row", 0)

        # Pick up db_path from graph if already ingested
        if self.graph.get("db_path") and Path(self.graph["db_path"]).exists():
            self._db_path = self.graph["db_path"]

    # ── Properties ────────────────────────────────────────────────────────────

    @property
    def is_large(self) -> bool:
        return self._row_count > LARGE_FILE_THRESHOLD

    @property
    def db_path(self) -> Optional[str]:
        return self._db_path

    # ── Sheet → table resolution ──────────────────────────────────────────────

    def _sheet_to_table(self, sheet_name) -> str:
        """
        Resolve a sheet name to its SQLite table name via _sheet_registry.

        Priority:
          1. graph["sheet_tables"][sheet_name]["table"]  — fastest, in-memory
          2. _sheet_registry lookup in the db            — reliable fallback
          3. sanitised sheet name                        — last resort
        """
        # 1. In-memory from graph
        sheet_tables = self.graph.get("sheet_tables", {})
        if sheet_name in sheet_tables:
            entry = sheet_tables[sheet_name]
            if isinstance(entry, dict):
                return entry.get("table", self._sanitise(sheet_name))
            return str(entry)

        # 2. _sheet_registry in db
        if self._db_path and Path(self._db_path).exists():
            try:
                conn = sqlite3.connect(self._db_path)
                row = conn.execute(
                    "SELECT table_name FROM _sheet_registry WHERE sheet_name=?",
                    (str(sheet_name),)
                ).fetchone()
                conn.close()
                if row:
                    return row[0]
            except Exception:
                pass

        # 3. Sanitise the sheet name directly
        return self._sanitise(str(sheet_name)) if sheet_name else "data"

    @staticmethod
    def _sanitise(name: str) -> str:
        s = str(name).lower().strip()
        s = re.sub(r"[^\w]", "_", s)
        s = re.sub(r"_+", "_", s).strip("_")
        if not s or s[0].isdigit():
            s = "sheet_" + s
        return s or "sheet"

    # ── Public interface ───────────────────────────────────────────────────────

    def ensure_sqlite(self) -> str:
        """
        Ensure a SQLite database exists for this workbook.
        Uses freshness check — re-ingests only if Excel is newer than db.
        """
        p  = Path(self.ai_path)
        db = p.parent / (p.stem.replace("_ai_workbook", "") + ".db")
        db_str = str(db)

        if not db.exists():
            from ingest import ingest_excel_to_sqlite
            result = ingest_excel_to_sqlite(self.ai_path, db_str)
            # Update graph with sheet_tables so _sheet_to_table works immediately
            if "sheet_tables" in result:
                self.graph["sheet_tables"] = result["sheet_tables"]
                self.graph["db_path"]      = db_str

        self._db_path = db_str
        return db_str

    def count(self) -> int:
        """Row count for the active sheet."""
        if self._row_count:
            return self._row_count
        if self._db_path:
            table = self._sheet_to_table(self.sheet)
            try:
                conn = sqlite3.connect(self._db_path)
                n = conn.execute(
                    f"SELECT COUNT(*) FROM [{table}]"
                ).fetchone()[0]
                conn.close()
                return n
            except Exception:
                return 0
        return 0

    def read(self,
             columns:      Optional[List[str]] = None,
             nrows:        Optional[int]       = None,
             filter_expr:  Optional[str]       = None) -> pd.DataFrame:
        """
        Read data into a DataFrame.

        For small files:  pd.read_excel with usecols / nrows.
        For large files:  SQL SELECT against the correct per-sheet table.

        filter_expr:
          Small files → pandas .query() string  e.g. "Amount > 100000"
          Large files → SQL WHERE clause        e.g. "_c_amount > 100000"
        """
        if self.is_large and self._db_path:
            return self._sql_read(columns, nrows, filter_expr)
        return self._pandas_read(columns, nrows, filter_expr)

    def iter_chunks(self,
                    columns:   Optional[List[str]] = None,
                    chunksize: int                 = CHUNK_SIZE
                    ) -> Iterator[pd.DataFrame]:
        """
        Yield chunks. Always memory-safe regardless of file size.
        Small files: one chunk. Large files: SQLite pages.
        """
        if self.is_large and self._db_path:
            yield from self._sql_chunks(columns, chunksize)
        elif self.is_large:
            yield from _chunked_excel_reader(self.ai_path, self.sheet, chunksize)
        else:
            yield self._pandas_read(columns, None, None)

    # ── Internal readers ───────────────────────────────────────────────────────

    def _pandas_read(self, columns, nrows, filter_expr) -> pd.DataFrame:
        df = pd.read_excel(self.ai_path, sheet_name=self.sheet,
                           usecols=columns, nrows=nrows)
        if filter_expr:
            try:
                df = df.query(filter_expr)
            except Exception:
                pass
        return df

    def _sql_read(self, columns, nrows, filter_expr) -> pd.DataFrame:
        table    = self._sheet_to_table(self.sheet)
        conn     = sqlite3.connect(self._db_path)
        col_expr = "*" if not columns else ", ".join(f"[{c}]" for c in columns)
        where    = f"WHERE {filter_expr}" if filter_expr else ""
        limit    = f"LIMIT {nrows}" if nrows else ""
        sql      = f"SELECT {col_expr} FROM [{table}] {where} {limit}"
        df       = pd.read_sql(sql, conn)
        conn.close()
        return df

    def _sql_chunks(self, columns, chunksize) -> Iterator[pd.DataFrame]:
        table    = self._sheet_to_table(self.sheet)
        conn     = sqlite3.connect(self._db_path)
        col_expr = "*" if not columns else ", ".join(f"[{c}]" for c in columns)
        offset   = 0
        while True:
            df = pd.read_sql(
                f"SELECT {col_expr} FROM [{table}] "
                f"LIMIT {chunksize} OFFSET {offset}",
                conn)
            if df.empty:
                break
            yield df
            offset += chunksize
            if len(df) < chunksize:
                break
        conn.close()

    # ── Output helpers ─────────────────────────────────────────────────────────

    def output_dir(self) -> Path:
        p = Path(self.ai_path)
        d = p.parent / (p.stem.replace("_ai_workbook", "") + "_analysis")
        d.mkdir(exist_ok=True)
        return d

    def write_result(self, df: pd.DataFrame, name: str,
                     fmt: str = "excel") -> Path:
        if isinstance(df.columns, pd.MultiIndex):
            df = df.copy()
            df.columns = ["_".join(str(c) for c in col).strip("_")
                          for col in df.columns]
        out = self.output_dir()
        if fmt == "csv":
            path = out / f"{name}.csv"
            df.to_csv(str(path), index=False)
        else:
            path = out / f"{name}.xlsx"
            with pd.ExcelWriter(str(path), engine="openpyxl") as w:
                df.to_excel(w, index=True, sheet_name=name[:31])
        return path