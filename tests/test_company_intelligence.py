from types import SimpleNamespace

from darwill_ai_prospector.services.company_intelligence import (
    build_company_intelligence,
    build_email_resolution,
)


def item(**kwargs):
    defaults = dict(
        company_name="Test Home Services",
        company_state="TX",
        company_city="Dallas",
        company_revenue=15000000,
        company_employees=75,
        why_company="Residential HVAC company with homeowner focus.",
        recommended_strategy="Lead with new movers.",
        company_research_summary="Residential home services hiring technicians.",
        company_deep_research="",
        contact_acquisition_report="",
        contact_email="",
        contact_source_type="",
        contact_email_status="",
        contact_email_verification_status="",
        contact_email_recovery_method="",
        contact_zoominfo_enrichment_result="",
        contact_email_confidence=0,
        contact_predicted_email_pattern="",
        contact_zoominfo_email_availability="Unknown",
    )
    defaults.update(kwargs)
    return SimpleNamespace(**defaults)


def test_company_fit():
    result = build_company_intelligence(item())
    assert result.confidence >= 70
    assert "residential" in result.residential_assessment.lower()


def test_limit_exceeded_email_state():
    result = build_email_resolution(
        item(
            contact_zoominfo_enrichment_result=(
                "ZoomInfo MCP enrichment limit exceeded"
            )
        )
    )
    assert result.status == "zoominfo_limit"
    assert "manual ZoomInfo export" in result.recommended_action
