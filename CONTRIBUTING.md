# Contributing to ByeByeExcel

This document outlines proposed upgrades and improvements for future consideration. Since this repository is primarily maintained for internal use, not all proposals will be implemented.

## Current State

- **Version**: 2.x (MCP-based architecture)
- **Core Focus**: Excel analysis with LLM integration
- **Key Strengths**:
  - Large file support (1M+ rows via SQLite chunking)
  - Domain-agnostic design
  - Zero external dependencies for knowledge search (SQLite FTS5)

---

## Proposed Upgrades

### 1. Fuzzy Deduplication Integration

**Status**: `bogus_detect.py` and `fuzzy_dedup.py` removed in v2 cleanup

**Proposal**: Re-integrate fuzzy deduplication as a native tool with:
- Block-based comparison (3-char keys)
- Configurable normalisers: Naga names, Indian names, generic text, numeric ranges, dates, codes
- Band thresholding (High/Review/Possible)
- HTML visualisation of dedup results

**Priority**: Medium - useful for data cleaning workflows

---

### 2. Enhanced Audit System

**Status**: `audit.py` present but basic implementation

**Proposal**: Expand audit capabilities to include:
- Formula propagation tracking (detect stale computed values)
- Reference validation (check for #REF! patterns in formulas)
- Data type consistency across sheets
- Cross-sheet dependency analysis

**Priority**: Low - current 8-issue audit suffices for most cases

---

### 3. PDF Knowledge Extraction Pipeline

**Status**: `corpus_builder.py` present, needs testing

**Proposal**: Complete the PDF → knowledge base pipeline:
- Document parsing with layout preservation
- Section-aware extraction prompts
- Entry validation and deduplication
- Bulk import from markdown/JSON files

**Priority**: Medium - enables documentation reuse across projects

---

### 4. Web UI for Analysis Results

**Status**: Not implemented

**Proposal**: Simple web interface to view:
- Audit results with issue filtering
- Dedup cluster visualisation
- Chart exports gallery
- Query history and execution logs

**Technical Approach**:
- FastAPI backend serving analysis data
- Chart.js for visualisations
- Local-only (no server deployment needed)

**Priority**: Low - CLI-based workflow preferred by current users

---

### 5. Performance Optimisations

**Status**: Basic chunked processing in place

**Proposal**:
- Parallel sheet processing where safe
- Incremental graph updates (avoid full re-scan after edits)
- Caching layer for repeated queries
- Query plan optimisation for complex tasks

**Priority**: High - would significantly improve responsiveness

---

### 6. Configuration System

**Status**: Hardcoded thresholds throughout codebase

**Proposal**: Central configuration file:
```json
{
  "fuzzy": {
    "high_threshold": 0.80,
    "review_threshold": 0.60,
    "block_chars": 3,
    "max_block_size": 150
  },
  "audit": {
    "skip_formula_override_check": false,
    "include_vba_analysis": true
  },
  "output": {
    "max_inline_rows": 40,
    "chart_dpi": 100
  }
}
```

**Priority**: Medium - improves user configurability

---

### 7. Version Control Integration

**Status**: Not implemented

**Proposal**: Track workbook changes:
- Git-style history of analysis operations
- Revert to previous states if needed
- Diff view for schema changes between versions

**Priority**: Low - niche requirement

---

## Testing Strategy

All upgrades should include:

1. **Unit Tests**: Per-module functionality
2. **Integration Tests**: End-to-end workflows
3. **Test Excel Files**: Various sizes and complexities
4. **Benchmark Suite**: Performance baselines

Current test files located in `test_excel/` (to be restored from backup).

---

## Contribution Guidelines

1. **Start with a Plan**: Propose changes via issue before coding
2. **Follow Existing Patterns**: Match naming, structure, and style
3. **Test Thoroughly**: Run against multiple Excel file types
4. **Document Changes**: Update README.md or DESIGN.md as needed
5. **Keep It Simple**: Avoid premature optimisation

---

## Reporting Issues

Include:
- Python version
- Excel file sample (if applicable)
- Error message and stack trace
- Expected vs actual behaviour

---

*Last updated: 2026*
