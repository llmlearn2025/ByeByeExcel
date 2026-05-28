"""
normalisers.py — Pluggable text normalisation strategies for fuzzy dedup.

The LLM reads the graph (column names, sample values, format_meaning) and
calls excel_get_fuzzy_options() to understand what normalisers exist.
The LLM then selects which normaliser(s) to apply per column.
The MCP executes — it never decides which normaliser fits.

Built-in normalisers:
  phonetic_names  — Aspirated consonant folding + clan/tribal suffix removal.
                    For names where the same person is recorded with phonetic
                    spelling variants across data entry operators.
  title_names     — Strips honorific titles and collapses double letters.
                    For names that appear with/without titles (Dr, Prof, etc.)
                    and common transliteration variants.
  us_names        — Western name normalisation. Strips titles (Mr/Mrs/Dr/Jr/Sr),
                    normalises Mc/Mac prefixes, hyphenated surnames.
  generic_text    — Lowercase + unicode strip + punctuation collapse only.
                    Safe default for any text column.
  address         — Strips flat/floor/building/house number prefixes.
  numeric_range   — Treats numbers within ±N% as equivalent.
  date_approx     — Treats dates within ±N days as equivalent.
  code_id         — Strips leading zeros, spaces, separators from ID codes.

Region configuration (config.py → COUNTRY):
  "IN"      → phonetic_names and title_names suggested for name columns
  "US"      → us_names suggested for name columns
  "GENERIC" → generic_text suggested for all name columns

Adding a new region normaliser: see README.md → "Extending Normalisers"

Each normaliser exposes:
  normalise(value: str) -> str         for single values
  score(a: str, b: str) -> float       0–1 composite match score
  describe() -> dict                   returned to LLM by excel_get_fuzzy_options
"""

import re, unicodedata, datetime
from typing import Optional
from rapidfuzz import fuzz

try:
    from config import COUNTRY
except ImportError:
    COUNTRY = "GENERIC"


# ─────────────────────────────────────────────────────────────────────────────
# BASE CLASS
# ─────────────────────────────────────────────────────────────────────────────

class BaseNormaliser:
    name: str = "base"
    description: str = ""
    best_for: list = []
    not_for: list = []

    def normalise(self, value) -> str:
        if value is None: return ""
        s = unicodedata.normalize("NFKD", str(value))
        s = "".join(c for c in s if not unicodedata.combining(c))
        return re.sub(r"\s+", " ", s.lower().strip())

    def score(self, a, b) -> float:
        na, nb = self.normalise(a), self.normalise(b)
        if not na or not nb: return 0.0
        if na == nb: return 1.0
        return max(
            fuzz.ratio(na, nb),
            fuzz.partial_ratio(na, nb),
            fuzz.token_sort_ratio(na, nb),
            fuzz.token_set_ratio(na, nb),
        ) / 100.0

    def describe(self) -> dict:
        return {
            "name":        self.name,
            "description": self.description,
            "best_for":    self.best_for,
            "not_for":     self.not_for,
            "example":     self._example(),
        }

    def _example(self) -> dict:
        return {}


# ─────────────────────────────────────────────────────────────────────────────
# PHONETIC NAMES
# For names with aspirated consonant variants and clan/tribal suffixes.
# Originally developed for South/Southeast Asian phonetic patterns but
# applies wherever names have aspirated vs unaspirated spelling variants.
# ─────────────────────────────────────────────────────────────────────────────

_PHONETIC_RULES = [
    # Aspirated → unaspirated consonant folding
    (r'\btshe\b', 'tse'), (r'\bkhe\b', 'ke'),  (r'\bphe\b', 'pe'),
    (r'\bche\b',  'tse'), (r'\bthe\b',  'te'),  (r'\bdhe\b', 'de'),
    (r'\bghe\b',  'ge'),  (r'\bjhe\b',  'je'),
    # ng/nk consonant confusion
    (r'nk(?=[aeiou])', 'ng'),
    # Double letter collapse
    (r'([bcdfghjklmnpqrstvwxyz])\1+', r'\1'),
    # Common suffix/ending variants
    (r'tham\b', 'tam'), (r'than\b', 'tan'), (r'them\b', 'tem'),
    (r'ham\b',  'am'),  (r'han\b',  'an'),
    # Common titles/honorifics (extend this list for your region — see README)
    (r'^(mr|mrs|ms|dr|prof|rev|late)\b\.?\s*', ''),
    # Clan / tribal / community suffixes that may appear or be omitted
    # across records for the same person (extend for your region — see README)
    (r'\b(ao|lotha|sumi|angami|chakhesang|zeliang|yimchunger|phom|'
     r'khiamniungan|sangtam|rengma|pochuri|konyak|chang|makware|'
     r'liangmai|rongmei|anal|tangkul|mao|maram|poumai|zeme)\b', ''),
]

class PhoneticsNamesNormaliser(BaseNormaliser):
    """
    Normaliser for names where:
    - The same person is spelled differently by different data entry operators
      due to aspirated vs unaspirated consonants (e.g. Tse/Tshe, Khe/Ke)
    - Names include a clan/tribal/community suffix that may or may not
      appear across records for the same person

    The _PHONETIC_RULES list is intentionally open for extension.
    See README.md → "Extending Normalisers" to add rules for your region.
    """
    name = "phonetic_names"
    description = (
        "Normalises names with aspirated consonant variants and optional "
        "clan/tribal/community suffixes. Folds Tse/Tshe, Khe/Ke variants, "
        "strips suffixes that appear inconsistently across records, "
        "collapses double letters and ending variants."
    )
    best_for = [
        "Name columns where the same name has multiple phonetic spellings",
        "Datasets with clan, tribal, or community suffixes in names",
        "Records entered by multiple operators with different phonetic conventions",
    ]
    not_for = [
        "Names without aspirated consonants or suffix patterns",
        "Numeric or date columns",
        "Address columns (use address normaliser)",
    ]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        s = re.sub(r"[^\w\s]", " ", s)
        for pattern, replacement in _PHONETIC_RULES:
            s = re.sub(pattern, replacement, s, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a":     "Tsering Norbu",
            "input_b":     "Tshering Norbu",
            "normalised_a": self.normalise("Tsering Norbu"),
            "normalised_b": self.normalise("Tshering Norbu"),
            "score":        round(self.score("Tsering Norbu", "Tshering Norbu"), 3),
        }


# Keep legacy alias so saved queries referencing "naga_names" continue to work
NagaNamesNormaliser = PhoneticsNamesNormaliser


# ─────────────────────────────────────────────────────────────────────────────
# TITLE NAMES (IN region default)
# For names with honorific titles and transliteration variants.
# ─────────────────────────────────────────────────────────────────────────────

_TITLE_RULES = [
    # Titles — extend for your region (see README)
    (r'^(mr|mrs|ms|dr|prof|col|brig|lt|capt|sgt|rev|late|'
     r'shri|smt|kumari|sri|srimati|pu|puni)\b\.?\s*', ''),
    # Common transliteration variants
    (r'([aeiou])\1+', r'\1'),              # double vowels
    (r'([bcdfghjklmnpqrstvwxyz])\1+', r'\1'),  # double consonants
]

class TitleNamesNormaliser(BaseNormaliser):
    """
    Normaliser for names that include honorific titles which may or may not
    appear across records for the same person. Also collapses double letters
    common in transliterated names.

    The _TITLE_RULES list is open for extension.
    See README.md → "Extending Normalisers".
    """
    name = "title_names"
    description = (
        "Strips honorific titles (Mr/Mrs/Dr/Prof and region-specific equivalents) "
        "and collapses double letters common in transliterated names. "
        "Less aggressive than phonetic_names — use when names are spelled "
        "consistently but titles appear inconsistently."
    )
    best_for = [
        "Name columns where titles appear in some records but not others",
        "Datasets with mixed formal/informal name entry",
        "Names with double-letter transliteration variants",
    ]
    not_for = [
        "Names with aspirated consonant variants (use phonetic_names)",
        "Non-text columns",
    ]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        s = re.sub(r"[^\w\s]", " ", s)
        for pattern, replacement in _TITLE_RULES:
            s = re.sub(pattern, replacement, s, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a": "Dr. Sarah Johnson",
            "input_b": "Sarah Johnson",
            "score":   round(self.score("Dr. Sarah Johnson", "Sarah Johnson"), 3),
        }


# Keep legacy alias
IndianNamesNormaliser = TitleNamesNormaliser


# ─────────────────────────────────────────────────────────────────────────────
# US NAMES
# Western name normalisation for US/UK/AU/CA datasets.
# ─────────────────────────────────────────────────────────────────────────────

_US_RULES = [
    # Titles and suffixes
    (r'^(mr|mrs|ms|miss|dr|prof|rev|hon|sir)\b\.?\s*', ''),
    (r'\b(jr|sr|ii|iii|iv|esq|phd|md|dds|dvm)\b\.?\s*$', ''),
    # Mc/Mac prefix normalisation (McDonald/MacDonald)
    (r'\bmac(?=[a-z])', 'mc'),
    # Hyphenated surnames — keep both parts
    (r'-', ' '),
    # Double letters
    (r'([bcdfghjklmnpqrstvwxyz])\1+', r'\1'),
]

class USNamesNormaliser(BaseNormaliser):
    """
    Normaliser for Western/US-style names.
    Handles Mr/Mrs/Dr/Jr/Sr titles, Mc/Mac prefix variants,
    hyphenated surnames, and double-letter collisions.

    The _US_RULES list is open for extension.
    See README.md → "Extending Normalisers".
    """
    name = "us_names"
    description = (
        "Western name normalisation. Strips titles (Mr/Mrs/Dr/Jr/Sr/II/III), "
        "normalises Mc/Mac prefix variants (McDonald = MacDonald), "
        "expands hyphenated surnames, collapses double letters."
    )
    best_for = [
        "Name columns in US, UK, Canadian, or Australian datasets",
        "Names with Western honorific titles",
        "Records with Mc/Mac prefix inconsistency",
        "Hyphenated surname matching",
    ]
    not_for = [
        "Names with aspirated consonant patterns (use phonetic_names)",
        "Names with non-Western titles",
        "Numeric or date columns",
    ]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        s = re.sub(r"[^\w\s-]", " ", s)   # keep hyphen for _US_RULES
        for pattern, replacement in _US_RULES:
            s = re.sub(pattern, replacement, s, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a": "Dr. John McDonald Jr.",
            "input_b": "John MacDonald",
            "score":   round(self.score("Dr. John McDonald Jr.", "John MacDonald"), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# GENERIC TEXT
# ─────────────────────────────────────────────────────────────────────────────

class GenericTextNormaliser(BaseNormaliser):
    name = "generic_text"
    description = (
        "Minimal normalisation: unicode strip, lowercase, punctuation collapse, "
        "whitespace collapse. Safe default for any text column when you don't "
        "know the domain. No phonetic rules — pure character edit distance."
    )
    best_for = [
        "Place or location names",
        "Category, scheme, or label fields",
        "Any text column where phonetic rules would be harmful",
        "First pass on unknown data",
    ]
    not_for = [
        "Person names with phonetic variants (use phonetic_names or title_names)",
        "Numeric columns",
    ]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        return re.sub(r"[^\w\s]", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a": "North Street Branch",
            "input_b": "north street  branch",
            "score":   round(self.score("North Street Branch",
                                        "north street  branch"), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# ADDRESS
# ─────────────────────────────────────────────────────────────────────────────

_ADDRESS_RULES = [
    (r'\b(flat|floor|f/f|g/f|house|h\.?no\.?|plot|door|ward|apt|suite|unit)\b'
     r'[\s\-]*\d*', ''),
    (r'\b(near|opp|opposite|beside|behind|above|below|next\s*to)\b', ''),
    (r'\b(street|st\.?|road|rd\.?|lane|ln\.?|avenue|ave\.?|drive|dr\.?|'
     r'boulevard|blvd\.?|nagar|colony|layout|extension|ext\.?|'
     r'phase|sector|block)\b', ''),
    (r'\d{4,6}',  ''),    # postal/zip codes (4-6 digit)
    (r'[,\./\\]', ' '),
    (r'\s+',      ' '),
]

class AddressNormaliser(BaseNormaliser):
    name = "address"
    description = (
        "Strips unit/flat numbers, floor numbers, and generic address words "
        "(Street, Road, Lane, Avenue, Drive, Colony) before comparison. "
        "Useful when the same address appears with/without a unit number."
    )
    best_for = [
        "Address columns in any dataset",
        "Vendor or contractor address deduplication",
    ]
    not_for = [
        "Person names",
        "When unit number differences are meaningful (e.g. delivery address)",
    ]

    def normalise(self, value) -> str:
        s = super().normalise(value)
        for pattern, replacement in _ADDRESS_RULES:
            s = re.sub(pattern, replacement, s, flags=re.IGNORECASE)
        return re.sub(r"\s+", " ", s).strip()

    def _example(self) -> dict:
        return {
            "input_a": "Apt 4B, 42 Main Street",
            "input_b": "42 Main Street",
            "score":   round(self.score("Apt 4B, 42 Main Street",
                                        "42 Main Street"), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# NUMERIC RANGE
# ─────────────────────────────────────────────────────────────────────────────

class NumericRangeNormaliser(BaseNormaliser):
    name = "numeric_range"
    description = (
        "Matches numeric values within a configurable tolerance. "
        "Numbers within ±tolerance% of each other score 1.0. "
        "Useful for amount columns with rounding or measurement columns."
    )
    best_for = [
        "Amount or payment columns with minor rounding differences",
        "Age columns where ±1 year is acceptable",
        "Quantity or measurement columns with tolerances",
    ]
    not_for = [
        "ID or code columns (use code_id)",
        "Name or text columns",
        "Date columns (use date_approx)",
    ]

    def __init__(self, tolerance: float = 0.05):
        self.tolerance = tolerance

    def normalise(self, value) -> str:
        try:
            return str(round(float(str(value).replace(",", "")), 2))
        except (ValueError, TypeError):
            return str(value or "")

    def score(self, a, b) -> float:
        try:
            fa = float(str(a).replace(",", ""))
            fb = float(str(b).replace(",", ""))
            if fa == 0 and fb == 0: return 1.0
            if fa == 0 or fb == 0: return 0.0
            pct_diff = abs(fa - fb) / max(abs(fa), abs(fb))
            if pct_diff <= self.tolerance: return 1.0
            if pct_diff >= 0.5: return 0.0
            return max(0.0, 1.0 - pct_diff / 0.5)
        except (ValueError, TypeError):
            return 0.0

    def _example(self) -> dict:
        return {
            "input_a":  "150000",
            "input_b":  "151500",
            "tolerance": self.tolerance,
            "score":    round(self.score(150000, 151500), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# DATE APPROX
# ─────────────────────────────────────────────────────────────────────────────

class DateApproxNormaliser(BaseNormaliser):
    name = "date_approx"
    description = (
        "Parses dates in various formats and scores based on how many days "
        "apart they are. Tolerance of 30 days handles transposed day/month, "
        "estimated dates, and data entry errors."
    )
    best_for = [
        "Date of birth columns (DOB often estimated or transposed)",
        "Registration or event date columns with ±few days tolerance",
    ]
    not_for = [
        "Name or text columns",
        "When exact date match is required (e.g. payment date audit)",
    ]

    _DATE_FMTS = [
        "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y",
        "%d-%m-%y", "%Y/%m/%d", "%d.%m.%Y", "%B %d, %Y",
        "%d %B %Y", "%b %d %Y", "%m/%d/%Y", "%m-%d-%Y",
    ]

    def __init__(self, day_tolerance: int = 30):
        self.day_tolerance = day_tolerance

    def _parse(self, value) -> Optional[datetime.date]:
        s = str(value).strip()
        for fmt in self._DATE_FMTS:
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

    def normalise(self, value) -> str:
        d = self._parse(value)
        return d.isoformat() if d else str(value or "")

    def score(self, a, b) -> float:
        da, db = self._parse(a), self._parse(b)
        if da is None or db is None:
            return fuzz.ratio(str(a), str(b)) / 100.0
        diff = abs((da - db).days)
        if diff == 0: return 1.0
        if diff >= self.day_tolerance * 3: return 0.0
        return max(0.0, 1.0 - diff / (self.day_tolerance * 3))

    def _example(self) -> dict:
        return {
            "input_a": "2024-03-15",
            "input_b": "15/03/2024",
            "note":    "same date, different format",
            "score":   round(self.score("2024-03-15", "15/03/2024"), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# CODE / ID
# ─────────────────────────────────────────────────────────────────────────────

class CodeIdNormaliser(BaseNormaliser):
    name = "code_id"
    description = (
        "Strips leading zeros, spaces, hyphens, and separators from ID codes. "
        "Useful for national ID numbers, registration numbers, or any code "
        "field where the same value appears with inconsistent formatting."
    )
    best_for = [
        "National or government ID number columns",
        "Registration or reference number columns",
        "Any ID/code column with formatting inconsistencies (spaces, dashes)",
    ]
    not_for = [
        "Name columns",
        "When leading zeros are semantically meaningful",
    ]

    def normalise(self, value) -> str:
        s = re.sub(r"[\s\-_./]", "", str(value or "").lower())
        return s.lstrip("0") or "0"

    def score(self, a, b) -> float:
        na, nb = self.normalise(a), self.normalise(b)
        if not na or not nb: return 0.0
        if na == nb: return 1.0
        return fuzz.ratio(na, nb) / 100.0

    def _example(self) -> dict:
        return {
            "input_a": "123-45-6789",
            "input_b": "123456789",
            "score":   round(self.score("123-45-6789", "123456789"), 3),
        }


# ─────────────────────────────────────────────────────────────────────────────
# REGISTRY
# ─────────────────────────────────────────────────────────────────────────────

REGISTRY: dict[str, BaseNormaliser] = {
    "phonetic_names": PhoneticsNamesNormaliser(),
    "title_names":    TitleNamesNormaliser(),
    "us_names":       USNamesNormaliser(),
    "generic_text":   GenericTextNormaliser(),
    "address":        AddressNormaliser(),
    "numeric_range":  NumericRangeNormaliser(),
    "date_approx":    DateApproxNormaliser(),
    "code_id":        CodeIdNormaliser(),
    # Legacy aliases — kept so saved queries using old names continue to work
    "naga_names":     PhoneticsNamesNormaliser(),
    "indian_names":   TitleNamesNormaliser(),
}


def get_normaliser(name: str, **kwargs) -> BaseNormaliser:
    """Retrieve a normaliser by name, with optional parameter overrides."""
    if name not in REGISTRY:
        raise ValueError(
            f"Unknown normaliser '{name}'. "
            f"Available: {[k for k in REGISTRY if not k in ('naga_names','indian_names')]}"
        )
    if name == "numeric_range" and "tolerance" in kwargs:
        return NumericRangeNormaliser(tolerance=kwargs["tolerance"])
    if name == "date_approx" and "day_tolerance" in kwargs:
        return DateApproxNormaliser(day_tolerance=kwargs["day_tolerance"])
    return REGISTRY[name]


# ─────────────────────────────────────────────────────────────────────────────
# SUGGESTION ENGINE
# Uses COUNTRY from config.py to rank normalisers for name columns.
# ─────────────────────────────────────────────────────────────────────────────

# Map COUNTRY code → preferred name normaliser + reason
_COUNTRY_NAME_NORMALISER = {
    "IN": {
        "primary":   "phonetic_names",
        "secondary": "title_names",
        "reason":    "phonetic_names handles aspirated consonant variants and "
                     "clan/community suffixes common in South Asian name data. "
                     "title_names handles honorific prefixes (Dr/Prof/Shri/Smt).",
    },
    "US": {
        "primary":   "us_names",
        "secondary": "generic_text",
        "reason":    "us_names handles Western titles (Mr/Mrs/Dr/Jr/Sr), "
                     "Mc/Mac prefix variants, and hyphenated surnames.",
    },
    "GENERIC": {
        "primary":   "generic_text",
        "secondary": None,
        "reason":    "No region configured. generic_text is the safe default. "
                     "Set COUNTRY in config.py for region-specific suggestions.",
    },
}


def suggest_for_column(col_name: str, sample_values: list,
                       format_meaning: str = "general",
                       fill_meaning: str = "none") -> list:
    """
    Given what the LLM knows about a column from the graph,
    return ranked normaliser suggestions with reasoning.
    The LLM uses this output to decide — not the MCP.

    Suggestions are influenced by COUNTRY in config.py for name columns.
    All normalisers are always available regardless of COUNTRY setting.
    """
    col_lower   = col_name.lower()
    suggestions = []
    region      = _COUNTRY_NAME_NORMALISER.get(
        COUNTRY, _COUNTRY_NAME_NORMALISER["GENERIC"])

    # ── Date columns ──────────────────────────────────────────────────────────
    if format_meaning == "date" or re.search(r"dob|birth|date", col_lower):
        suggestions.append({
            "normaliser": "date_approx",
            "confidence": "high",
            "reason":     f"Column '{col_name}' has date format or date-like name. "
                          "date_approx handles format variants and transposed day/month.",
        })

    # ── Numeric / amount columns ───────────────────────────────────────────────
    if format_meaning in ("currency_inr", "currency_usd", "currency",
                          "number_thousands", "percentage"):
        suggestions.append({
            "normaliser": "numeric_range",
            "confidence": "high",
            "reason":     f"Column '{col_name}' has numeric format '{format_meaning}'. "
                          "numeric_range handles rounding differences.",
        })

    # ── ID / code columns ─────────────────────────────────────────────────────
    if re.search(r"(\bid\b|code|no\b|num|ref|ssn|sin|nino|passport|license)",
                 col_lower):
        suggestions.append({
            "normaliser": "code_id",
            "confidence": "high",
            "reason":     f"Column '{col_name}' looks like an ID or code field. "
                          "code_id strips formatting separators and leading zeros.",
        })

    # ── Name columns — use region config ──────────────────────────────────────
    if re.search(r"name|naam|apellido|nom\b|nombre", col_lower):
        suggestions.append({
            "normaliser": region["primary"],
            "confidence": "medium",
            "reason":     f"Column '{col_name}' is a name column. "
                          f"{region['reason']}",
        })
        if region["secondary"]:
            suggestions.append({
                "normaliser": region["secondary"],
                "confidence": "low",
                "reason":     f"Alternative if {region['primary']} is too aggressive "
                              f"for your data.",
            })

    # ── Address / location columns ─────────────────────────────────────────────
    if re.search(r"address|addr|street|locality|suburb|postcode|zipcode",
                 col_lower):
        suggestions.append({
            "normaliser": "address",
            "confidence": "medium",
            "reason":     f"Column '{col_name}' is an address field. "
                          "address strips unit numbers and generic street words.",
        })
    elif re.search(r"city|town|village|district|county|state|region|location",
                   col_lower):
        suggestions.append({
            "normaliser": "generic_text",
            "confidence": "medium",
            "reason":     f"Column '{col_name}' is a place name field. "
                          "generic_text normalises case and whitespace.",
        })

    # ── Default fallback ───────────────────────────────────────────────────────
    if not suggestions:
        suggestions.append({
            "normaliser": "generic_text",
            "confidence": "low",
            "reason":     f"No specific pattern detected for '{col_name}'. "
                          "generic_text is the safe default. Review column samples "
                          "and choose a more specific normaliser if appropriate.",
        })

    return suggestions


def describe_all() -> dict:
    """Return full description of all normalisers for LLM context."""
    # Exclude legacy aliases from description
    return {
        name: n.describe()
        for name, n in REGISTRY.items()
        if name not in ("naga_names", "indian_names")
    }