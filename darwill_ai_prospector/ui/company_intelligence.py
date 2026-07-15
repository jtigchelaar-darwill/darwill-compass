"""Presentation-only company intelligence for the Compass Deal Desk."""

from __future__ import annotations

import re
from typing import Any


RESIDENTIAL_TERMS = {
    "residential", "homeowner", "home service", "home services",
    "homes", "household", "homeowners",
}
GROWTH_TERMS = {
    "hiring": ("hiring", "careers", "job openings", "now hiring"),
    "expansion": ("expansion", "new location", "expanded", "service area"),
    "acquisition": ("acquired", "acquisition", "merged", "private equity"),
    "marketing investment": (
        "google ads", "paid search", "direct mail", "advertising",
        "marketing", "membership", "financing",
    ),
}
TECH_TERMS = {
    "ServiceTitan": ("servicetitan",),
    "Housecall Pro": ("housecall pro", "housecallpro"),
    "HubSpot": ("hubspot",),
    "Salesforce": ("salesforce",),
    "Google Analytics": ("google analytics", "gtag"),
    "CallRail": ("callrail",),
    "Podium": ("podium",),
    "Birdeye": ("birdeye",),
}


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _contains_any(text: str, terms: tuple[str, ...] | set[str]) -> bool:
    lowered = text.lower()
    return any(term in lowered for term in terms)


def _detected_labels(
    text: str,
    library: dict[str, tuple[str, ...]],
) -> list[str]:
    return [
        label for label, terms in library.items()
        if _contains_any(text, terms)
    ]


def _score_label(score: int) -> str:
    if score >= 90:
        return f"{score}% · Excellent"
    if score >= 75:
        return f"{score}% · Strong"
    if score >= 55:
        return f"{score}% · Moderate"
    return f"{score}% · Limited"


def build_company_intelligence(item: Any) -> dict[str, str]:
    """Summarize stored qualification evidence without changing engine logic."""
    why = _clean(getattr(item, "why_company", ""))
    evidence = _clean(getattr(item, "evidence", ""))
    sources = _clean(getattr(item, "sources", ""))
    website = _clean(getattr(item, "company_website", ""))
    city = _clean(getattr(item, "company_city", ""))
    state = _clean(getattr(item, "company_state", ""))
    revenue = getattr(item, "company_revenue", None)
    employees = getattr(item, "company_employees", None)
    strategy = _clean(getattr(item, "recommended_strategy", ""))
    outreach_confidence = int(
        getattr(item, "outreach_confidence", 0) or 0
    )

    combined = " ".join([why, evidence, sources, strategy]).strip()
    lowered = combined.lower()

    # Evidence completeness—not a replacement for the discovery fit score.
    evidence_score = 38
    evidence_score += 18 if why else 0
    evidence_score += 14 if evidence else 0
    evidence_score += 8 if sources else 0
    evidence_score += 8 if revenue else 0
    evidence_score += 6 if employees else 0
    evidence_score += 6 if state else 0
    evidence_score = min(99, evidence_score)

    residential_hits = sum(
        1 for term in RESIDENTIAL_TERMS if term in lowered
    )
    residential_score = min(
        99,
        42
        + residential_hits * 13
        + (10 if "residential" in lowered else 0)
        + (7 if "homeowner" in lowered else 0),
    )
    if not combined:
        residential_score = 0

    if revenue and employees:
        size_label = f"${revenue / 1_000_000:.1f}M · {employees:,} employees"
    elif revenue:
        size_label = f"${revenue / 1_000_000:.1f}M revenue"
    elif employees:
        size_label = f"{employees:,} employees"
    else:
        size_label = "Size unavailable"

    growth = _detected_labels(combined, GROWTH_TERMS)
    technologies = _detected_labels(combined, TECH_TERMS)

    service_area = "No stored service-area evidence"
    area_match = re.search(
        r"((?:service|serving|territor(?:y|ies)|market(?:s)?)"
        r".{0,140})",
        combined,
        flags=re.I,
    )
    if area_match:
        service_area = area_match.group(1).strip(" .|")
    elif city or state:
        service_area = ", ".join(part for part in [city, state] if part)

    opportunity_score = min(
        99,
        max(
            45,
            round(
                outreach_confidence * 0.65
                + evidence_score * 0.20
                + residential_score * 0.15
            ),
        ),
    )

    if strategy:
        darwill_angle = (
            f"Lead with {strategy}. "
            f"{why or 'Use the stored company evidence to connect the offer to a specific growth opportunity.'}"
        )
    elif growth:
        darwill_angle = (
            "Lead with measurable customer acquisition and market growth. "
            f"Stored evidence indicates {', '.join(growth)}."
        )
    else:
        darwill_angle = (
            "Lead with Darwill's new-mover and measurable customer-acquisition "
            "capabilities, using the qualification evidence below as the personalized hook."
        )

    evidence_lines = [
        f"COMPANY\n{_clean(getattr(item, 'company_name', ''))}",
        f"WHY QUALIFIED\n{why or 'No qualification explanation was stored.'}",
        f"STORED EVIDENCE\n{evidence or 'No additional evidence was stored.'}",
        f"RECOMMENDED STRATEGY\n{strategy or 'No strategy was stored.'}",
        f"SOURCES\n{sources or website or 'No source URL was stored.'}",
        (
            "IMPORTANT\n"
            "The percentages above summarize evidence completeness for review. "
            "They do not replace or alter the proven discovery engine's qualification decision."
        ),
    ]

    return {
        "qualification": _score_label(evidence_score),
        "residential": _score_label(residential_score),
        "size": size_label,
        "opportunity": _score_label(opportunity_score),
        "location": ", ".join(part for part in [city, state] if part) or "Unavailable",
        "website": website or "Unavailable",
        "growth": ", ".join(growth) if growth else "No stored growth or marketing signal",
        "technology": ", ".join(technologies) if technologies else "No stored technology signal",
        "service_area": service_area,
        "darwill_angle": darwill_angle,
        "evidence_text": "\n\n".join(evidence_lines),
    }
