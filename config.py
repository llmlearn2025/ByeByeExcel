"""
config.py — Central configuration for ByeByeExcel MCP server.

Edit this file to configure:
  - Server port
  - Country/region (controls which name normalisers are active by default)
  - Storage thresholds
  - Output preferences

The country setting determines which region-specific normaliser the system
suggests first when it encounters name columns. It does NOT disable any
normaliser — all normalisers remain available regardless of country.
The LLM always makes the final choice.
"""

from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# SERVER
# ─────────────────────────────────────────────────────────────────────────────

PORT = 6699

# ─────────────────────────────────────────────────────────────────────────────
# REGION
# Controls the default name normaliser suggested for name columns.
# Does NOT disable other normalisers — the LLM can always override.
#
# Supported values:
#   "IN"   — India. Suggests phonetic_names (aspirated consonants, clan suffixes)
#             and title_names (Shri/Smt/Dr) for name columns.
#   "US"   — United States. Suggests us_names (Western name titles, hyphenated
#             surnames, Mc/Mac prefix normalisation) for name columns.
#   "GENERIC" — No region-specific suggestion. Always suggests generic_text
#               for name columns. Use this when your data is international
#               or you prefer to select normalisers manually.
#
# To add a new region: see the "Adding a new region" section in README.md
# ─────────────────────────────────────────────────────────────────────────────

COUNTRY = "IN"   # Change to "US" or "GENERIC" as needed

# ─────────────────────────────────────────────────────────────────────────────
# STORAGE
# ─────────────────────────────────────────────────────────────────────────────

# Files above this row count use SQLite instead of in-memory pandas.
# Lower this if you have limited RAM (e.g. 16GB → 30000).
# Raise this if you have plenty of RAM and want faster reads (e.g. 100000).
LARGE_FILE_THRESHOLD = 50_000

# Rows per chunk when streaming large files.
CHUNK_SIZE = 10_000

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

# Results with this many rows or fewer are returned as inline markdown tables.
# Above this: saved to HTML file or Excel export.
MAX_INLINE_ROWS = 40

# DPI for generated charts.
CHART_DPI = 120

# ─────────────────────────────────────────────────────────────────────────────
# KNOWLEDGE BASE
# ─────────────────────────────────────────────────────────────────────────────

# Path to the shared SQLite knowledge base file.
# Default: same directory as config.py
KB_PATH = str(Path(__file__).parent / "excel_knowledge.db")