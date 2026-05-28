"""
knowledge.py — Documentation search + workbook query memory for Excel MCP v2.

Two subsystems:

─────────────────────────────────────────────────────────────────
SUBSYSTEM 1: KNOWLEDGE BASE (doc + example search)
─────────────────────────────────────────────────────────────────

A SQLite FTS5 knowledge base storing documentation entries and
working code examples from authoritative sources (Excel, pandas,
SQL, data analysis textbooks).

Structure of one entry:
  {
    "id":         "pandas-groupby-001",
    "category":   "pandas|excel|sql|analysis|visualisation",
    "subcategory":"groupby|pivot|string|date|window|chart|...",
    "title":      "Rolling 12-month average by category",
    "problem":    "Plain English description of what this solves",
    "tags":       ["rolling", "window", "average", "time-series"],
    "code":       "Python/pandas/SQL code that actually works",
    "notes":      "Gotchas, edge cases, when NOT to use this",
    "source":     "Book/doc title + chapter",
    "difficulty": "beginner|intermediate|advanced",
  }

FTS5 gives BM25 ranking natively — no embedding model needed.
Handles 1 lakh entries at millisecond speed.

Query flow:
  1. LLM calls excel_search_docs("rolling average by month", top_k=10)
  2. MCP queries FTS5 on title + problem + tags, returns ranked entries
  3. LLM reads entries, uses the code as a template
  4. If LLM adapts and succeeds, it calls excel_save_query() to persist

─────────────────────────────────────────────────────────────────
SUBSYSTEM 2: WORKBOOK QUERY MEMORY
─────────────────────────────────────────────────────────────────

Per-workbook JSON file: <stem>_queries.json
Each entry records a successful query with:
  - description (natural language)
  - columns involved (from the graph at time of execution)
  - sheet name
  - code that was executed
  - result summary
  - timestamp
  - execution_count (incremented each time this query is reused)

On every excel_search_docs call, the MCP also searches the
workbook's own query memory for matches. Working queries for
THIS workbook are ranked above generic documentation.

Write-back:
  LLM calls excel_save_query() after any successful analysis.
  The MCP deduplicates by semantic key (description + sheet + cols).
  High-frequency queries bubble up in future searches.
"""

import re, json, sqlite3, hashlib, datetime
from pathlib import Path
from typing import Optional

# ─────────────────────────────────────────────────────────────────────────────
# DEFAULT KNOWLEDGE DB PATH
# ─────────────────────────────────────────────────────────────────────────────

# The knowledge base is shared across all workbooks.
# Place it in the same directory as the MCP server files.
_DEFAULT_KB_PATH = Path(__file__).parent / "excel_knowledge.db"


# ─────────────────────────────────────────────────────────────────────────────
# KNOWLEDGE BASE — SCHEMA + INIT
# ─────────────────────────────────────────────────────────────────────────────

def init_knowledge_db(db_path: str = None) -> str:
    """
    Create the knowledge base SQLite file if it doesn't exist.
    Returns the db path.
    Schema:
      docs        — FTS5 virtual table (title, problem, tags, code, notes)
      docs_meta   — metadata for docs (id, category, subcategory, source, difficulty)
      doc_stats   — access stats per entry (for ranking boost)
    """
    path = db_path or str(_DEFAULT_KB_PATH)
    conn = sqlite3.connect(path)
    cur  = conn.cursor()

    # FTS5 virtual table for full-text search
    cur.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS docs USING fts5(
            entry_id,
            category,
            subcategory,
            title,
            problem,
            tags,
            code,
            notes,
            tokenize = 'porter unicode61'
        )
    """)

    # Metadata (non-FTS fields)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS docs_meta (
            entry_id    TEXT PRIMARY KEY,
            source      TEXT,
            difficulty  TEXT,
            created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Access stats (used to boost frequently-useful entries)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS doc_stats (
            entry_id        TEXT PRIMARY KEY,
            access_count    INTEGER DEFAULT 0,
            last_accessed   DATETIME,
            helpful_votes   INTEGER DEFAULT 0
        )
    """)

    conn.commit()
    conn.close()
    return path


def insert_entry(entry: dict, db_path: str = None) -> str:
    """
    Insert one knowledge entry. Skips if entry_id already exists.
    Returns entry_id.

    entry fields:
      id, category, subcategory, title, problem, tags (list or str),
      code, notes, source, difficulty
    """
    path = db_path or str(_DEFAULT_KB_PATH)
    conn = sqlite3.connect(path)
    cur  = conn.cursor()

    eid  = entry.get("id") or _make_id(entry.get("title",""))
    tags = entry.get("tags", [])
    if isinstance(tags, list):
        tags = " ".join(tags)

    # Skip if exists
    cur.execute("SELECT entry_id FROM docs WHERE entry_id=?", (eid,))
    if cur.fetchone():
        conn.close()
        return eid

    cur.execute("""
        INSERT INTO docs(entry_id, category, subcategory, title,
                         problem, tags, code, notes)
        VALUES (?,?,?,?,?,?,?,?)
    """, (
        eid,
        entry.get("category","general"),
        entry.get("subcategory",""),
        entry.get("title",""),
        entry.get("problem",""),
        tags,
        entry.get("code",""),
        entry.get("notes",""),
    ))

    cur.execute("""
        INSERT OR IGNORE INTO docs_meta(entry_id, source, difficulty)
        VALUES (?,?,?)
    """, (eid, entry.get("source",""), entry.get("difficulty","intermediate")))

    cur.execute("""
        INSERT OR IGNORE INTO doc_stats(entry_id, access_count)
        VALUES (?,0)
    """, (eid,))

    conn.commit()
    conn.close()
    return eid


def bulk_insert(entries: list, db_path: str = None) -> dict:
    """Insert multiple entries. Returns {inserted, skipped, total}."""
    path = db_path or str(_DEFAULT_KB_PATH)
    conn = sqlite3.connect(path)
    cur  = conn.cursor()

    inserted = skipped = 0
    for entry in entries:
        eid  = entry.get("id") or _make_id(entry.get("title",""))
        cur.execute("SELECT entry_id FROM docs WHERE entry_id=?", (eid,))
        if cur.fetchone():
            skipped += 1
            continue
        tags = entry.get("tags",[])
        if isinstance(tags, list): tags = " ".join(tags)
        cur.execute("""
            INSERT INTO docs(entry_id, category, subcategory, title,
                             problem, tags, code, notes)
            VALUES (?,?,?,?,?,?,?,?)
        """, (eid, entry.get("category","general"),
              entry.get("subcategory",""), entry.get("title",""),
              entry.get("problem",""), tags,
              entry.get("code",""), entry.get("notes","")))
        cur.execute("""
            INSERT OR IGNORE INTO docs_meta(entry_id, source, difficulty)
            VALUES (?,?,?)
        """, (eid, entry.get("source",""), entry.get("difficulty","intermediate")))
        cur.execute("""
            INSERT OR IGNORE INTO doc_stats(entry_id, access_count) VALUES (?,0)
        """, (eid,))
        inserted += 1

    conn.commit()
    conn.close()
    return {"inserted": inserted, "skipped": skipped, "total": len(entries)}


def _make_id(title: str) -> str:
    return hashlib.md5(title.lower().strip().encode()).hexdigest()[:12]


# ─────────────────────────────────────────────────────────────────────────────
# SEARCH
# ─────────────────────────────────────────────────────────────────────────────

def search_docs(
    query: str,
    category: Optional[str] = None,
    difficulty: Optional[str] = None,
    top_k: int = 10,
    db_path: str = None,
) -> list:
    """
    FTS5 BM25 search over the knowledge base.
    Returns ranked list of entry dicts.

    query:      Natural language or keyword query
    category:   Optional filter: 'pandas'|'excel'|'sql'|'analysis'|'visualisation'
    difficulty: Optional filter: 'beginner'|'intermediate'|'advanced'
    top_k:      Number of results to return

    Ranking: FTS5 BM25 score + access_count boost (popular entries rank higher)
    """
    path = db_path or str(_DEFAULT_KB_PATH)
    if not Path(path).exists():
        return []

    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    # Build WHERE clause
    filters = []
    params  = []

    # Category filter (on FTS column — use MATCH ... AND category=...)
    # FTS5 allows filtering with WHERE on non-indexed columns via a join
    cat_filter   = "AND d.category = ?" if category  else ""
    diff_filter  = "AND m.difficulty = ?" if difficulty else ""
    if category:   params.append(category)
    if difficulty: params.append(difficulty)

    # Clean query for FTS5 (escape special chars)
    fts_query = _clean_fts_query(query)

    try:
        sql = f"""
            SELECT
                d.entry_id, d.category, d.subcategory, d.title,
                d.problem, d.tags, d.code, d.notes,
                m.source, m.difficulty,
                COALESCE(s.access_count, 0) as access_count,
                bm25(docs) as bm25_score
            FROM docs d
            JOIN docs_meta m ON d.entry_id = m.entry_id
            LEFT JOIN doc_stats s ON d.entry_id = s.entry_id
            WHERE docs MATCH ?
            {cat_filter}
            {diff_filter}
            ORDER BY (bm25_score - COALESCE(s.access_count, 0) * 0.01) ASC
            LIMIT ?
        """
        # BM25 in FTS5 returns negative values — lower is better
        all_params = [fts_query] + params + [top_k]
        cur.execute(sql, all_params)
        rows = cur.fetchall()
    except sqlite3.OperationalError:
        # FTS5 query syntax error — fall back to LIKE search
        rows = _fallback_search(cur, query, category, top_k)

    conn.close()

    results = []
    for row in rows:
        results.append({
            "id":          row["entry_id"],
            "category":    row["category"],
            "subcategory": row["subcategory"],
            "title":       row["title"],
            "problem":     row["problem"],
            "tags":        row["tags"],
            "code":        row["code"],
            "notes":       row["notes"],
            "source":      row["source"],
            "difficulty":  row["difficulty"],
            "access_count":row["access_count"],
        })
    return results


def _clean_fts_query(q: str) -> str:
    """Remove FTS5 special characters that cause syntax errors."""
    q = re.sub(r'["\*\^\(\)]', ' ', q)
    q = re.sub(r'\s+', ' ', q).strip()
    # Wrap multi-word in quotes for phrase search
    words = q.split()
    if len(words) == 1:
        return words[0]
    # FTS5: search for all words (implicit AND)
    return " ".join(words)


def _fallback_search(cur, query: str, category: str, top_k: int) -> list:
    """LIKE-based fallback when FTS5 query fails."""
    words = query.lower().split()[:3]
    like = "%" + "%".join(words) + "%"
    cat_filter = "AND category = ?" if category else ""
    params = [like, like, like]
    if category: params.append(category)
    params.append(top_k)
    cur.execute(f"""
        SELECT d.*, m.source, m.difficulty, COALESCE(s.access_count,0) as access_count,
               0 as bm25_score
        FROM docs d
        JOIN docs_meta m ON d.entry_id = m.entry_id
        LEFT JOIN doc_stats s ON d.entry_id = s.entry_id
        WHERE (lower(title) LIKE ? OR lower(problem) LIKE ? OR lower(tags) LIKE ?)
        {cat_filter}
        LIMIT ?
    """, params)
    return cur.fetchall()


def record_access(entry_id: str, helpful: bool = False, db_path: str = None):
    """Increment access counter for an entry (called when LLM uses it)."""
    path = db_path or str(_DEFAULT_KB_PATH)
    if not Path(path).exists(): return
    conn = sqlite3.connect(path)
    conn.execute("""
        UPDATE doc_stats
        SET access_count = access_count + 1,
            helpful_votes = helpful_votes + ?,
            last_accessed = datetime('now')
        WHERE entry_id = ?
    """, (1 if helpful else 0, entry_id))
    conn.commit()
    conn.close()


def db_stats(db_path: str = None) -> dict:
    """Return stats about the knowledge base."""
    path = db_path or str(_DEFAULT_KB_PATH)
    if not Path(path).exists():
        return {"exists": False, "path": path}
    conn = sqlite3.connect(path)
    total   = conn.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
    by_cat  = dict(conn.execute(
        "SELECT category, COUNT(*) FROM docs GROUP BY category").fetchall())
    by_diff = dict(conn.execute(
        "SELECT m.difficulty, COUNT(*) FROM docs d "
        "JOIN docs_meta m ON d.entry_id=m.entry_id "
        "GROUP BY m.difficulty").fetchall())
    top_accessed = conn.execute(
        "SELECT d.title, s.access_count FROM doc_stats s "
        "JOIN docs d ON s.entry_id=d.entry_id "
        "WHERE s.access_count > 0 ORDER BY s.access_count DESC LIMIT 10"
    ).fetchall()
    conn.close()
    return {
        "exists":       True,
        "path":         path,
        "total_entries":total,
        "by_category":  by_cat,
        "by_difficulty":by_diff,
        "top_accessed": [{"title": r[0], "count": r[1]} for r in top_accessed],
    }


# ─────────────────────────────────────────────────────────────────────────────
# WORKBOOK QUERY MEMORY
# ─────────────────────────────────────────────────────────────────────────────

def _query_memory_path(ai_path: str) -> Path:
    p = Path(ai_path).resolve()
    return p.parent / (p.stem.replace("_ai_workbook","") + "_queries.json")


def load_query_memory(ai_path: str) -> dict:
    """Load the workbook's query memory. Returns {} if none exists."""
    path = _query_memory_path(ai_path)
    if not path.exists():
        return {"workbook": str(ai_path), "queries": [], "version": "1"}
    try:
        return json.loads(path.read_text())
    except Exception:
        return {"workbook": str(ai_path), "queries": [], "version": "1"}


def save_query(
    ai_path: str,
    description: str,
    code: str,
    sheet: str = "",
    columns_used: list = None,
    result_summary: str = "",
    tags: list = None,
) -> dict:
    """
    Persist a successful query to the workbook's query memory.
    Deduplicates by semantic key (description + sheet + sorted columns).
    Increments execution_count if same query already exists.
    Returns the saved entry.
    """
    mem = load_query_memory(ai_path)

    cols   = sorted(columns_used or [])
    sem_key = hashlib.md5(
        f"{description.lower().strip()}|{sheet}|{','.join(cols)}".encode()
    ).hexdigest()[:12]

    # Check for existing entry with same semantic key
    for entry in mem["queries"]:
        if entry.get("key") == sem_key:
            entry["execution_count"] = entry.get("execution_count", 1) + 1
            entry["last_used"]       = datetime.datetime.now().isoformat()
            entry["code"]            = code  # update to latest version
            if result_summary:
                entry["result_summary"] = result_summary
            _flush_memory(ai_path, mem)
            return entry

    # New entry
    entry = {
        "key":             sem_key,
        "description":     description,
        "sheet":           sheet,
        "columns_used":    cols,
        "tags":            tags or [],
        "code":            code,
        "result_summary":  result_summary,
        "execution_count": 1,
        "created_at":      datetime.datetime.now().isoformat(),
        "last_used":       datetime.datetime.now().isoformat(),
    }
    mem["queries"].append(entry)
    _flush_memory(ai_path, mem)
    return entry


def search_query_memory(
    ai_path: str,
    query: str,
    sheet: str = "",
    top_k: int = 5,
) -> list:
    """
    Search the workbook's query memory for similar past queries.
    Uses simple keyword overlap scoring — no external model needed.
    Returns ranked list of matching entries.
    """
    mem = load_query_memory(ai_path)
    if not mem["queries"]:
        return []

    q_words = set(re.sub(r"[^\w\s]","",query.lower()).split())
    scored  = []

    for entry in mem["queries"]:
        d_words = set(re.sub(r"[^\w\s]","",
                             entry["description"].lower()).split())
        t_words = set(w.lower() for w in entry.get("tags",[]))
        c_words = set(c.lower().replace("_"," ")
                      for c in entry.get("columns_used",[]))
        all_words = d_words | t_words | c_words

        # Jaccard-like overlap
        overlap = len(q_words & all_words)
        if overlap == 0:
            continue

        score = overlap / max(len(q_words | all_words), 1)
        # Boost by execution count (proven queries rank higher)
        score += entry.get("execution_count", 1) * 0.02
        # Boost if sheet matches
        if sheet and entry.get("sheet") == sheet:
            score += 0.1

        scored.append((score, entry))

    scored.sort(key=lambda x: -x[0])
    return [e for _, e in scored[:top_k]]


def _flush_memory(ai_path: str, mem: dict):
    path = _query_memory_path(ai_path)
    path.write_text(json.dumps(mem, indent=2, ensure_ascii=False))


def get_memory_stats(ai_path: str) -> dict:
    """Summary of workbook query memory for inclusion in graph."""
    mem = load_query_memory(ai_path)
    queries = mem.get("queries", [])
    if not queries:
        return {"count": 0}
    top = sorted(queries, key=lambda x: -x.get("execution_count", 1))[:5]
    return {
        "count":       len(queries),
        "top_queries": [{"desc": e["description"],
                         "runs": e.get("execution_count",1)}
                        for e in top],
    }


# ─────────────────────────────────────────────────────────────────────────────
# RENDER — for LLM consumption
# ─────────────────────────────────────────────────────────────────────────────

def render_search_results(
    doc_results: list,
    memory_results: list,
    query: str,
    ai_path: str = "",
) -> str:
    """
    Render search results as structured markdown for LLM.
    Workbook memory results come first (proven, specific to this file).
    Doc results follow (general knowledge).
    """
    L = []; a = L.append
    a(f"## Knowledge Search: *{query}*\n")

    # ── Workbook memory first ─────────────────────────────────────────────────
    if memory_results:
        a("### From workbook memory (previously working code for this file)")
        a(f"*These queries ran successfully on this workbook before.*\n")
        for i, e in enumerate(memory_results, 1):
            cols_str = ", ".join(f"`{c}`" for c in e.get("columns_used",[]))
            a(f"#### [{i}] {e['description']}")
            if e.get("sheet"):     a(f"- **Sheet:** `{e['sheet']}`")
            if cols_str:           a(f"- **Columns:** {cols_str}")
            a(f"- **Run count:** {e.get('execution_count',1)}")
            if e.get("result_summary"):
                a(f"- **Last result:** {e['result_summary'][:100]}")
            a(f"```python\n{e['code'].strip()}\n```")
            a("")
    else:
        if ai_path:
            a("*No matching queries in workbook memory yet. "
              "Successful queries will be saved here automatically.*\n")

    # ── Doc results ───────────────────────────────────────────────────────────
    if doc_results:
        a("### From knowledge base")
        a(f"*{len(doc_results)} results ranked by relevance.*\n")
        for i, e in enumerate(doc_results, 1):
            tags_str = e.get("tags","").replace(" ",", ")
            a(f"#### [{i}] {e['title']}")
            a(f"*{e.get('category','')} / {e.get('subcategory','')} "
              f"— {e.get('difficulty','')}*")
            if e.get("problem"):    a(f"**Problem:** {e['problem']}")
            if tags_str:            a(f"**Tags:** {tags_str}")
            if e.get("source"):     a(f"**Source:** {e['source']}")
            if e.get("code"):
                a(f"```python\n{e['code'].strip()}\n```")
            if e.get("notes"):
                a(f"> **Note:** {e['notes']}")
            a("")
    elif not memory_results:
        a("*No matching entries in knowledge base.*")
        a("The corpus may not cover this topic yet.")
        a("See `example_doc.md` for instructions on adding entries.")

    return "\n".join(L)
