import re
from typing import Literal

CommandComplexity = Literal["simple", "compound_explicit", "semantic_complex"]

# Keywords that indicate the command needs interpretation or content generation.
# NOTE: "full table", "header row", etc. are NOT listed here because the planner
# can resolve them deterministically using tracked context (last_table_range,
# header_range). Only include phrases that truly require LLM content generation
# or that the deterministic layer cannot handle at all.
_SEMANTIC_PHRASES = (
    "neatly", "professional", "professionally", "clean", "cleanly",
    "look good", "make it look", "looks good",
    "realistic", "sample data", "realistic data", "fake data",
    "summarize", "summarise", "analyze", "analyse",
    "about ", "report about", "presentation about", "document about",
    "based on", "relevant", "appropriate", "suitable",
    "make it", "style it", "stylish", "attractive", "elegant",
    "generate content", "add content",
)

# Conjunctions / list words that split a command into multiple clauses.
_CLAUSE_STARTERS = re.compile(
    r"""
    \bafter\s+that\b |
    \bfinally\b      |
    \bfollowed\s+by\b |
    \bnext\b         |
    \bthen\b         |
    \balso\b         |
    ;\s*             |
    \n
    """,
    re.VERBOSE | re.IGNORECASE,
)

# Simple comma splits that precede a new action verb (like the planner uses).
_COMMA_ACTION_RE = re.compile(
    r",\s+(?=(?:create|make|add|insert|write|set|apply|bold|italic|underline|"
    r"border|rename|save|open|protect|unprotect|close|autofit|on\s+slide)\b)",
    re.IGNORECASE,
)


def _strip_quoted(text: str) -> str:
    """Remove quoted substrings so their contents don't confuse heuristics."""
    return re.sub(r"""(['"])[^'"]{1,200}\1""", " ", text)


def classify_office_command_complexity(command: str) -> CommandComplexity:
    """
    Classify a natural-language Office command into one of three complexity buckets:

    - "simple"           — one intent, one app, explicit params, no ambiguity
    - "compound_explicit"— multiple direct clauses joined by conjunctions
    - "semantic_complex" — requires interpretation, content generation, or inferred ranges
    """
    text = (command or "").strip()
    if not text:
        return "simple"

    low = _strip_quoted(text).lower()

    # ── Semantic-complex signals ────────────────────────────────────────────
    for phrase in _SEMANTIC_PHRASES:
        if phrase in low:
            return "semantic_complex"

    # ── Multi-clause detection ──────────────────────────────────────────────
    has_clause_starter = bool(_CLAUSE_STARTERS.search(low))
    has_comma_action   = bool(_COMMA_ACTION_RE.search(text))

    if has_clause_starter or has_comma_action:
        return "compound_explicit"

    # ── Simple: single clause, no vague language ────────────────────────────
    return "simple"
