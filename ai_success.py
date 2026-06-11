"""
AI Project Success Scoring
--------------------------
Companion module for the Closure Document Generator.

Design principles
=================
1. **Deterministic where it should be.** Budget variance and schedule
   adherence are pure arithmetic. We compute those sub-scores in Python so
   they are reproducible and never hallucinated. The LLM is only asked to
   judge the *qualitative* criteria (objectives met, deliverables completed,
   risk handling, execution quality).
2. **Structured output.** The model is instructed to return STRICT JSON.
   We parse it defensively and fall back to a heuristic if anything is off.
3. **Runs anywhere.** The OpenAI key is read from st.secrets first, then
   the environment — so a local .env and Streamlit Cloud Secrets both work
   with no code change.
4. **Never breaks the app.** Any failure (no key, no network, bad JSON)
   returns a structured result with engine="fallback" instead of raising.

Public API
==========
    assess_project_success(payload, weights=None, model="gpt-4o-mini") -> dict

`payload` keys used (all optional, missing -> treated as empty):
    project_name, overview, objectives, deliverables, risks, lessons,
    status, budget {planned, actual, currency},
    start_date (ISO str|date), end_date (ISO str|date)

Return shape
============
    {
      "final_score": 78.4,                 # 0-100, rounded to 1 dp
      "classification": "Successful",
      "band_emoji": "🟢",
      "criteria": [
         {"key","label","weight","score","justification"},
         ...
      ],
      "engine": "openai" | "fallback",
      "model": "gpt-4o-mini",
      "notes": "...",                      # optional human-readable note
    }
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date, datetime


# --------------------------------------------------------------------------- #
# Scoring model definition
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Criterion:
    key: str
    label: str
    default_weight: float   # 0..1, the six default weights sum to 1.0
    llm_scored: bool        # True -> GPT judges it; False -> computed in Python


CRITERIA: list[Criterion] = [
    Criterion("objectives",   "Objectives Achievement",     0.30, llm_scored=True),
    Criterion("deliverables", "Deliverables Completion",    0.25, llm_scored=True),
    Criterion("budget",       "Budget Adherence",           0.15, llm_scored=False),
    Criterion("schedule",     "Schedule Adherence",         0.10, llm_scored=False),
    Criterion("risk",         "Risk Management",            0.10, llm_scored=True),
    Criterion("quality",      "Stakeholder & Quality",      0.10, llm_scored=True),
]

DEFAULT_WEIGHTS: dict[str, float] = {c.key: c.default_weight for c in CRITERIA}
LLM_CRITERIA = [c for c in CRITERIA if c.llm_scored]


# Classification bands: (min_inclusive, label, emoji)
BANDS = [
    (85.0, "Highly Successful", "🟢"),
    (70.0, "Successful",        "🟢"),
    (50.0, "Partially Successful", "🟡"),
    (30.0, "Marginal / At Risk",   "🟠"),
    (0.0,  "Unsuccessful",      "🔴"),
]


def classify(score: float) -> tuple[str, str]:
    """Map a 0-100 score to (classification, emoji)."""
    for threshold, label, emoji in BANDS:
        if score >= threshold:
            return label, emoji
    return "Unsuccessful", "🔴"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _get_api_key() -> str | None:
    """st.secrets first (Streamlit Cloud), then environment (local .env)."""
    try:
        import streamlit as st  # imported lazily so the module is testable headless
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:  # noqa: BLE001 - secrets file absent locally is fine
        pass
    return os.environ.get("OPENAI_API_KEY")


def _to_date(value) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, datetime):
        return value.date()
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _clamp(x: float, lo: float = 0.0, hi: float = 100.0) -> float:
    return max(lo, min(hi, x))


def normalize_weights(weights: dict[str, float] | None) -> dict[str, float]:
    """Accept partial / unnormalized weights; return weights that sum to 1.0.
    Falls back to defaults if the input is empty or sums to zero."""
    if not weights:
        return dict(DEFAULT_WEIGHTS)
    cleaned = {c.key: max(0.0, float(weights.get(c.key, DEFAULT_WEIGHTS[c.key])))
               for c in CRITERIA}
    total = sum(cleaned.values())
    if total <= 0:
        return dict(DEFAULT_WEIGHTS)
    return {k: v / total for k, v in cleaned.items()}


# --------------------------------------------------------------------------- #
# Deterministic sub-scores (no LLM)
# --------------------------------------------------------------------------- #
def _budget_score(budget: dict | None) -> tuple[float, str]:
    """100 at/under budget; degrades as overspend % grows. Underspend is not
    penalised (treated as on-target) but flagged if very large."""
    if not budget:
        return 60.0, "No budget data provided — neutral score applied."
    planned = float(budget.get("planned") or 0)
    actual = float(budget.get("actual") or 0)
    if planned <= 0:
        return 60.0, "Planned budget missing — variance can't be computed."
    variance_pct = (actual - planned) / planned * 100.0  # +ve = overspend
    if variance_pct <= 0:
        # Under or on budget. Large underspend can signal poor estimation.
        under = -variance_pct
        if under > 25:
            return 80.0, f"Significantly under budget ({under:.1f}%) — possible over-estimation."
        return 100.0, f"On or under budget ({-variance_pct:+.1f}%)."
    # Overspend: lose ~2.5 points per 1% over, floored at 0.
    score = _clamp(100.0 - variance_pct * 2.5)
    return score, f"Over budget by {variance_pct:.1f}%."


def _schedule_score(start, end, status: str) -> tuple[float, str]:
    """We rarely have a *planned* end date separate from actual, so schedule
    is inferred from status + whether dates are coherent. This is intentionally
    conservative: it nudges, it doesn't dominate (10% weight)."""
    start_d, end_d = _to_date(start), _to_date(end)
    status_l = (status or "").lower()

    if "cancelled" in status_l:
        return 10.0, "Project cancelled."
    if "hold" in status_l:
        return 40.0, "Project on hold."

    base = 100.0 if "deviation" not in status_l else 75.0
    note = "Completed on plan." if base == 100.0 else "Completed with deviations."

    if start_d and end_d:
        if end_d < start_d:
            return 30.0, "End date precedes start date — timeline inconsistent."
        note += f" Duration {(end_d - start_d).days} days."
    else:
        note += " (Dates incomplete.)"
    return base, note


# --------------------------------------------------------------------------- #
# LLM scoring of qualitative criteria
# --------------------------------------------------------------------------- #
def _build_prompt(payload: dict) -> str:
    budget = payload.get("budget") or {}
    return f"""You are a senior project assurance reviewer. Evaluate the project
closure information below and score ONLY these four qualitative criteria, each
from 0 to 100 (100 = fully achieved / excellent):

- objectives  : How fully were the stated Objectives achieved, judged against
                the Overview and Deliverables?
- deliverables: Were the promised Deliverables actually produced and complete?
- risk        : How well were Risks & Issues identified and mitigated?
- quality     : Overall execution quality, stakeholder handling, and the
                maturity of Lessons Learned.

Be critical and evidence-based. If a section is empty or vague, score it lower
and say so. Do NOT score budget or schedule (handled separately).

PROJECT NAME: {payload.get('project_name','(unnamed)')}
STATUS: {payload.get('status','(unknown)')}

OVERVIEW:
{payload.get('overview','') or '(none)'}

OBJECTIVES:
{payload.get('objectives','') or '(none)'}

DELIVERABLES:
{payload.get('deliverables','') or '(none)'}

RISKS & ISSUES:
{payload.get('risks','') or '(none)'}

LESSONS LEARNED:
{payload.get('lessons','') or '(none)'}

BUDGET CONTEXT (for awareness only, do not score): planned={budget.get('planned')}, actual={budget.get('actual')}

Respond with STRICT JSON ONLY, no markdown, in exactly this shape:
{{
  "objectives":   {{"score": <int 0-100>, "justification": "<= 25 words"}},
  "deliverables": {{"score": <int 0-100>, "justification": "<= 25 words"}},
  "risk":         {{"score": <int 0-100>, "justification": "<= 25 words"}},
  "quality":      {{"score": <int 0-100>, "justification": "<= 25 words"}}
}}"""


def _call_openai(payload: dict, model: str) -> dict[str, dict] | None:
    """Returns {key: {score, justification}} for LLM criteria, or None on any
    failure (caller falls back to heuristic)."""
    api_key = _get_api_key()
    if not api_key:
        return None
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system",
                 "content": "You are a precise evaluator that returns only valid JSON."},
                {"role": "user", "content": _build_prompt(payload)},
            ],
            temperature=0.2,
            response_format={"type": "json_object"},
        )
        raw = resp.choices[0].message.content
        data = json.loads(raw)
        out = {}
        for c in LLM_CRITERIA:
            item = data.get(c.key, {}) or {}
            out[c.key] = {
                "score": _clamp(float(item.get("score", 50))),
                "justification": str(item.get("justification", "")).strip()
                                 or "No justification returned.",
            }
        return out
    except Exception as exc:  # noqa: BLE001 - degrade gracefully, never crash UI
        return {"__error__": {"score": 0, "justification": f"{type(exc).__name__}: {exc}"}}


def _heuristic_llm_scores(payload: dict) -> dict[str, dict]:
    """Offline fallback: score qualitative criteria by content richness so the
    feature still produces something sensible without an API call."""
    def richness(text: str) -> float:
        text = (text or "").strip()
        if not text:
            return 30.0
        lines = [l for l in text.splitlines() if l.strip()]
        words = len(text.split())
        score = 50 + min(40, words * 0.6) + min(10, len(lines) * 3)
        return _clamp(score)

    return {
        "objectives":   {"score": richness(payload.get("objectives")),
                         "justification": "Heuristic estimate from objectives detail (no AI key)."},
        "deliverables": {"score": richness(payload.get("deliverables")),
                         "justification": "Heuristic estimate from deliverables detail (no AI key)."},
        "risk":         {"score": richness(payload.get("risks")),
                         "justification": "Heuristic estimate from risk detail (no AI key)."},
        "quality":      {"score": richness(payload.get("lessons")),
                         "justification": "Heuristic estimate from lessons detail (no AI key)."},
    }


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #
def assess_project_success(payload: dict,
                           weights: dict[str, float] | None = None,
                           model: str = "gpt-4o-mini") -> dict:
    """Compute the weighted project success score. See module docstring."""
    weights = normalize_weights(weights)

    # 1. Deterministic sub-scores
    budget_score, budget_note = _budget_score(payload.get("budget"))
    schedule_score, schedule_note = _schedule_score(
        payload.get("start_date"), payload.get("end_date"), payload.get("status", "")
    )

    # 2. Qualitative sub-scores (LLM or fallback)
    llm = _call_openai(payload, model)
    engine = "openai"
    note = ""
    if llm is None:
        engine = "fallback"
        note = "No OpenAI API key found — used offline heuristic scoring."
        llm = _heuristic_llm_scores(payload)
    elif "__error__" in llm:
        engine = "fallback"
        note = f"AI call failed ({llm['__error__']['justification']}) — used heuristic scoring."
        llm = _heuristic_llm_scores(payload)

    computed = {
        "budget": {"score": budget_score, "justification": budget_note},
        "schedule": {"score": schedule_score, "justification": schedule_note},
    }

    # 3. Assemble per-criterion rows and weighted total
    criteria_out = []
    final = 0.0
    for c in CRITERIA:
        source = computed.get(c.key) or llm.get(c.key) or {"score": 50, "justification": ""}
        score = round(_clamp(float(source["score"])), 1)
        w = weights[c.key]
        final += score * w
        criteria_out.append({
            "key": c.key,
            "label": c.label,
            "weight": round(w, 4),
            "score": score,
            "justification": source["justification"],
        })

    final = round(final, 1)
    classification, emoji = classify(final)

    return {
        "final_score": final,
        "classification": classification,
        "band_emoji": emoji,
        "criteria": criteria_out,
        "engine": engine,
        "model": model if engine == "openai" else "heuristic",
        "notes": note,
        "assessed_at": datetime.now().isoformat(timespec="seconds"),
    }


# Quick manual test: python ai_success.py
if __name__ == "__main__":
    demo = {
        "project_name": "Migration of On-Prem Email to Cloud",
        "status": "Completed with deviations",
        "overview": "Migrated 2500+ users to cloud, improving availability.",
        "objectives": "- Reduce downtime < 1 hour\n- Improve remote access\n- Decommission legacy infra",
        "deliverables": "- Mailboxes migrated\n- Documentation created\n- Dashboard deployed",
        "risks": "Data corruption risk during cutover — mitigated with staged validation.",
        "lessons": "Validate data earlier; add a planning buffer for UAT.",
        "budget": {"planned": 5_000_000, "actual": 5_400_000, "currency": "₹ (INR)"},
        "start_date": "2025-01-06",
        "end_date": "2025-03-28",
    }
    import pprint
    pprint.pprint(assess_project_success(demo))
