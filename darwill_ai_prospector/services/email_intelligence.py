"""Email acquisition scoring and ZoomInfo credit recommendations."""

from __future__ import annotations
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class EmailIntelligenceDecision:
    public_status: str
    pattern_label: str
    confidence: int
    zoominfo_status: str
    recommendation: str
    reason: str
    credit_recommended: bool
    expected_gain: str


def _value(item: Any, name: str, default: Any = "") -> Any:
    return getattr(item, name, default)


def build_email_intelligence(item: Any) -> EmailIntelligenceDecision:
    email = str(_value(item, "contact_email", "") or "").strip()
    fields = [
        _value(item, "contact_source_type", ""),
        _value(item, "contact_email_status", ""),
        _value(item, "contact_email_verification_status", ""),
        _value(item, "contact_email_recovery_method", ""),
        _value(item, "contact_data_status", ""),
        _value(item, "contact_acquisition_report", ""),
    ]
    combined = " ".join(str(value or "") for value in fields).lower()
    explicit_availability = str(
        _value(item, "contact_zoominfo_email_availability", "Unknown")
        or "Unknown"
    ).strip()
    explicit_detail = str(
        _value(item, "contact_zoominfo_email_availability_detail", "") or ""
    ).strip()
    enrichment_attempted = bool(
        _value(item, "contact_zoominfo_enrichment_attempted", False)
    )
    enrichment_result = str(
        _value(item, "contact_zoominfo_enrichment_result", "") or ""
    ).strip()
    pattern = str(
        _value(item, "contact_predicted_email_pattern", "") or ""
    ).strip()
    support = int(_value(item, "contact_pattern_support_count", 0) or 0)
    current = int(_value(item, "contact_email_confidence", 0) or 0)
    contact_id = str(_value(item, "contact_id", "") or "")

    public = bool(email) and (
        contact_id.startswith("public-")
        or "company website" in combined
        or "public research" in combined
        or "press release" in combined
    ) and "zoominfo" not in combined

    if public:
        return EmailIntelligenceDecision(
            "Found free", pattern or "Not needed",
            min(100, max(current, 95)), "Not needed",
            "Use public email",
            "A public email is already available. Do not spend a credit.",
            False, "None — public email already present",
        )

    if email and "zoominfo" in combined:
        return EmailIntelligenceDecision(
            "Not public", pattern or "Not needed",
            min(100, max(current, 98)), "Email recovered",
            "Use recovered email",
            "ZoomInfo already supplied the email. No additional credit is needed.",
            False, "Already realized",
        )

    predicted = bool(email) and (
        "predicted" in combined or "pattern" in combined
    )
    if predicted:
        confidence = max(current, min(92, 55 + support * 12))
        return EmailIntelligenceDecision(
            "Predicted only", pattern or "Pattern inferred",
            confidence, "Check availability",
            "Validate before send" if confidence >= 75 else "Research further",
            "A predicted address exists but is not verified. Attempt free validation first.",
            False, "Potential verification upgrade",
        )

    if explicit_availability == "Available" and not email:
        return EmailIntelligenceDecision(
            "Not found", pattern or "No reliable pattern",
            max(current, min(90, 40 + support * 14)),
            "Verified email indicated",
            (
                "Review enrichment result"
                if enrichment_attempted
                else "Use 1 credit"
            ),
            (
                explicit_detail
                + (" " + enrichment_result if enrichment_result else "")
            ).strip(),
            not enrichment_attempted,
            "Missing → verified business email",
        )

    if explicit_availability == "Unavailable" and not email:
        return EmailIntelligenceDecision(
            "Not found", pattern or "No reliable pattern",
            max(current, min(90, 40 + support * 14)),
            "No email indicated", "Skip credit",
            explicit_detail or "ZoomInfo indicates no email is available.",
            False, "None expected",
        )

    available = any(phrase in combined for phrase in [
        "verified email available",
        "zoominfo email available",
        "email available in zoominfo",
        "enrichment email available",
    ])
    unavailable = any(phrase in combined for phrase in [
        "zoominfo unavailable",
        "no zoominfo email",
        "email unavailable",
        "no email available",
    ])
    confidence = max(current, min(90, 40 + support * 14))

    if available:
        return EmailIntelligenceDecision(
            "Not found", pattern or "No reliable pattern",
            confidence, "Verified email indicated", "Use 1 credit",
            "Public research failed and stored ZoomInfo evidence indicates a verified email is available.",
            True, "Missing → verified business email",
        )

    if unavailable:
        return EmailIntelligenceDecision(
            "Not found", pattern or "No reliable pattern",
            confidence, "No email indicated", "Skip credit",
            "ZoomInfo does not indicate an email is available, so a credit is unlikely to help.",
            False, "None expected",
        )

    return EmailIntelligenceDecision(
        "Not found", pattern or "No reliable pattern",
        confidence, "Availability unknown", "Check before credit",
        "Confirm that ZoomInfo can return an email before consuming a credit.",
        False, "Unknown until availability is confirmed",
    )


def format_email_intelligence_report(decision: EmailIntelligenceDecision) -> str:
    return (
        "EMAIL ACQUISITION DECISION\n\n"
        f"Public email status: {decision.public_status}\n"
        f"Pattern: {decision.pattern_label}\n"
        f"Current confidence: {decision.confidence}%\n"
        f"ZoomInfo status: {decision.zoominfo_status}\n"
        f"Expected gain: {decision.expected_gain}\n"
        f"Credit recommended: {'YES' if decision.credit_recommended else 'NO'}\n"
        f"Recommendation: {decision.recommendation}\n\n"
        f"Reason: {decision.reason}\n\n"
        "Policy: exhaust public/free sources first. Only use a ZoomInfo credit "
        "when evidence indicates that a verified email is available."
    )
