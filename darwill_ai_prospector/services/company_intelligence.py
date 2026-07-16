"""Explainable company and email intelligence for Deal Desk."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


def _get(item: Any, name: str, default: Any = "") -> Any:
    return getattr(item, name, default)


def _clean(value: Any) -> str:
    return str(value or "").strip()


@dataclass(frozen=True)
class EmailResolution:
    status: str
    label: str
    explanation: str
    recommended_action: str
    confidence: int


@dataclass(frozen=True)
class CompanyIntelligence:
    company_name: str
    fit_summary: str
    residential_assessment: str
    ownership_assessment: str
    geography_assessment: str
    scale_assessment: str
    growth_assessment: str
    marketing_assessment: str
    opportunity_assessment: str
    risk_assessment: str
    recommended_angle: str
    confidence: int


def build_email_resolution(item: Any) -> EmailResolution:
    email = _clean(_get(item, "contact_email"))
    source = " ".join(
        [
            _clean(_get(item, "contact_source_type")),
            _clean(_get(item, "contact_email_status")),
            _clean(_get(item, "contact_email_verification_status")),
            _clean(_get(item, "contact_email_recovery_method")),
            _clean(_get(item, "contact_zoominfo_enrichment_result")),
            _clean(_get(item, "contact_acquisition_report")),
        ]
    ).lower()
    confidence = int(_get(item, "contact_email_confidence", 0) or 0)
    predicted_pattern = _clean(
        _get(item, "contact_predicted_email_pattern")
    )
    availability = _clean(
        _get(item, "contact_zoominfo_email_availability", "Unknown")
    )
    enrichment_result = _clean(
        _get(item, "contact_zoominfo_enrichment_result")
    ).lower()

    if email and (
        "public" in source
        or "website" in source
        or "press release" in source
    ) and "zoominfo" not in source:
        return EmailResolution(
            status="public_verified",
            label="Public email found",
            explanation=(
                f"{email} was recovered from a public source. "
                "No ZoomInfo enrichment is needed."
            ),
            recommended_action="Use the public email",
            confidence=max(confidence, 95),
        )

    if email and "zoominfo" in source:
        return EmailResolution(
            status="zoominfo_verified",
            label="ZoomInfo email recovered",
            explanation=(
                f"{email} was returned by ZoomInfo enrichment."
            ),
            recommended_action="Use the verified ZoomInfo email",
            confidence=max(confidence, 98),
        )

    if email:
        return EmailResolution(
            status="email_available",
            label="Email available",
            explanation=(
                f"{email} is stored on the contact record, but its source "
                "classification is incomplete."
            ),
            recommended_action="Review source, then use",
            confidence=max(confidence, 75),
        )

    if "limit exceeded" in enrichment_result or "limit exceeded" in source:
        return EmailResolution(
            status="zoominfo_limit",
            label="ZoomInfo enrichment blocked",
            explanation=(
                "ZoomInfo MCP reported that the enrichment limit was exceeded. "
                "The contact may still have an email in ZoomInfo Sales, but "
                "Compass cannot retrieve it through MCP right now."
            ),
            recommended_action="Use manual ZoomInfo export or public research",
            confidence=100,
        )

    if availability == "Available":
        return EmailResolution(
            status="zoominfo_available",
            label="ZoomInfo says email exists",
            explanation=(
                "ZoomInfo search metadata indicates an email is available, "
                "but enrichment has not returned it."
            ),
            recommended_action="Retry after ZoomInfo resolves MCP limits",
            confidence=95,
        )

    if predicted_pattern:
        return EmailResolution(
            status="predicted",
            label="Predicted email only",
            explanation=(
                f"Compass inferred the pattern {predicted_pattern}, but no "
                "verified address was recovered."
            ),
            recommended_action="Validate before outreach",
            confidence=max(confidence, 50),
        )

    if availability == "Unavailable":
        return EmailResolution(
            status="unavailable",
            label="No ZoomInfo email indicated",
            explanation=(
                "Public research found no email and ZoomInfo metadata indicates "
                "that no email is available."
            ),
            recommended_action="Use phone, LinkedIn, or another contact",
            confidence=90,
        )

    return EmailResolution(
        status="missing",
        label="Email unresolved",
        explanation=(
            "No public email was found and ZoomInfo availability is unknown."
        ),
        recommended_action="Continue public research or choose another contact",
        confidence=40,
    )


def build_company_intelligence(item: Any) -> CompanyIntelligence:
    name = _clean(_get(item, "company_name")) or "Unknown company"
    state = _clean(_get(item, "company_state"))
    city = _clean(_get(item, "company_city"))
    revenue = int(_get(item, "company_revenue", 0) or 0)
    employees = int(_get(item, "company_employees", 0) or 0)
    why_company = _clean(_get(item, "why_company"))
    strategy = _clean(_get(item, "recommended_strategy"))
    research = " ".join(
        [
            _clean(_get(item, "company_research_summary")),
            _clean(_get(item, "company_deep_research")),
            _clean(_get(item, "why_company")),
            _clean(_get(item, "recommended_strategy")),
            _clean(_get(item, "contact_acquisition_report")),
        ]
    ).lower()

    residential_positive = any(
        phrase in research
        for phrase in [
            "residential",
            "homeowner",
            "home service",
            "home services",
            "homes",
        ]
    )
    commercial_risk = any(
        phrase in research
        for phrase in [
            "commercial-only",
            "commercial only",
            "primarily commercial",
            "industrial",
        ]
    )
    franchise_risk = any(
        phrase in research
        for phrase in ["franchise", "franchisee"]
    )
    pe_risk = any(
        phrase in research
        for phrase in [
            "private equity",
            "pe-backed",
            "portfolio company",
        ]
    )
    growth_signal = any(
        phrase in research
        for phrase in [
            "hiring",
            "expansion",
            "new location",
            "acquisition",
            "growing",
            "growth",
        ]
    )
    marketing_signal = any(
        phrase in research
        for phrase in [
            "marketing director",
            "digital marketing",
            "google ads",
            "meta",
            "hubspot",
            "servicetitan",
            "callrail",
        ]
    )

    if residential_positive and not commercial_risk:
        residential = "Strong evidence of residential/homeowner focus."
    elif commercial_risk:
        residential = "Commercial concentration requires manual review."
    else:
        residential = "Residential focus is not fully documented."

    if franchise_risk:
        ownership = "Possible franchise relationship; verify independence."
    elif pe_risk:
        ownership = "Possible private-equity ownership; verify eligibility."
    else:
        ownership = "No franchise or PE warning detected in stored research."

    location = ", ".join(part for part in [city, state] if part)
    geography = (
        f"Single listed market: {location}."
        if location
        else "Geographic footprint requires verification."
    )

    if revenue:
        scale = f"Estimated annual revenue: ${revenue:,.0f}."
    elif employees:
        scale = f"Employee count: approximately {employees:,}."
    else:
        scale = "Scale evidence is incomplete."

    growth = (
        "Stored research contains active growth or expansion signals."
        if growth_signal
        else "No strong recent growth trigger is stored."
    )
    marketing = (
        "Evidence suggests an established marketing or technology function."
        if marketing_signal
        else "Marketing sophistication is not yet strongly evidenced."
    )

    opportunity = (
        why_company
        or "Potential Darwill fit depends on residential mix, market scale, "
        "and access to a customer-acquisition decision maker."
    )
    recommended_angle = (
        strategy
        or "Lead with measurable homeowner acquisition and incremental growth."
    )

    risks = []
    if commercial_risk:
        risks.append("commercial concentration")
    if franchise_risk:
        risks.append("franchise status")
    if pe_risk:
        risks.append("private-equity ownership")
    risk = (
        "Review required: " + ", ".join(risks) + "."
        if risks
        else "No major stored disqualifier detected."
    )

    confidence = 45
    confidence += 20 if residential_positive else 0
    confidence += 10 if revenue or employees else 0
    confidence += 10 if why_company else 0
    confidence += 10 if strategy else 0
    confidence += 5 if growth_signal else 0
    confidence -= 20 if commercial_risk else 0
    confidence -= 10 if franchise_risk or pe_risk else 0
    confidence = max(0, min(100, confidence))

    fit_summary = (
        f"{name} appears suitable for further review."
        if confidence >= 70
        else f"{name} still requires qualification evidence."
    )

    return CompanyIntelligence(
        company_name=name,
        fit_summary=fit_summary,
        residential_assessment=residential,
        ownership_assessment=ownership,
        geography_assessment=geography,
        scale_assessment=scale,
        growth_assessment=growth,
        marketing_assessment=marketing,
        opportunity_assessment=opportunity,
        risk_assessment=risk,
        recommended_angle=recommended_angle,
        confidence=confidence,
    )


def format_company_intelligence(
    intelligence: CompanyIntelligence,
    email: EmailResolution,
) -> str:
    return (
        "COMPANY INTELLIGENCE\n\n"
        f"Company: {intelligence.company_name}\n"
        f"Confidence: {intelligence.confidence}%\n"
        f"Summary: {intelligence.fit_summary}\n\n"
        f"Residential fit: {intelligence.residential_assessment}\n"
        f"Ownership: {intelligence.ownership_assessment}\n"
        f"Geography: {intelligence.geography_assessment}\n"
        f"Scale: {intelligence.scale_assessment}\n"
        f"Growth: {intelligence.growth_assessment}\n"
        f"Marketing maturity: {intelligence.marketing_assessment}\n"
        f"Opportunity: {intelligence.opportunity_assessment}\n"
        f"Risk: {intelligence.risk_assessment}\n"
        f"Recommended angle: {intelligence.recommended_angle}\n\n"
        "EMAIL RESOLUTION\n\n"
        f"Status: {email.label}\n"
        f"Confidence: {email.confidence}%\n"
        f"Explanation: {email.explanation}\n"
        f"Recommended action: {email.recommended_action}"
    )
