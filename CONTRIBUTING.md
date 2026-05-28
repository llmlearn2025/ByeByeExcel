# Contributing to ByeByeExcel

This document outlines proposed upgrades and improvements for future consideration. Since this repository is primarily maintained for generic use and knowledge dissemination, not all proposals will be implemented. Users may directly fork and edit for their purpose

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
- Configurable normalisers: Names in local context, Indian names, generic text, numeric ranges, dates, codes
- Band thresholding (High/Review/Possible)
- HTML visualisation of dedup results

**Priority**: Medium - useful for data cleaning workflows
**Observation** : LLM performs better fuzzy deduplication using run_code than user provide code/claude generated dedupe. Use run_code and guide llm on what kind of dedup or fuzzy logic is needed. 
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
**The actual_ingestion.py has ability to bulk ingest from multiple files including pdf and epub. The code is added to repo along with code.py**
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
***Use run_code instead, Chart exports are working well**
---

### 5. Performance Optimisations

**Status**: Basic chunked processing in place

**Proposal**:
- Parallel sheet processing where safe
- Incremental graph updates (avoid full re-scan after edits)
- Caching layer for repeated queries
- Query plan optimisation for complex tasks

**Priority**: High - would significantly improve responsiveness

**excel update based db population is there. 50K rows take hardly a minute, hence Dropped**
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
**use run_code, same results available through llm - tested on gemma 4 26 B in LM studio. Results are better and use save skills to save the working code to the skill associated with the workbook**
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
**tested on 50K data and folder removed. Users can generate own test scripts if needed**
---

## Contribution & Forking Guidelines

This repository is intended as a static reference and knowledge base. The owner does not intend to maintain active version tracking, review pull requests, or release new updates. 

Instead, users are highly encouraged to **fork the repository and adapt the code directly** for their own purposes. If you choose to modify your fork, we recommend following these best practices:

1. **Keep It Simple**: Maintain the core philosophy of avoiding premature optimization.
2. **Follow Existing Patterns**: Match the naming conventions, structure, and style already present in the codebase.
3. **Test Thoroughly**: Run your changes against multiple Excel file types and sizes.
4. **Document Your Fork**: Update your fork's `README.md` to reflect your specific changes or custom workflows.

## Reporting Issues

Include:
- Python version
- Excel file sample (if applicable)
- Error message and stack trace
- Expected vs actual behaviour

---

*Last updated: 2026*
