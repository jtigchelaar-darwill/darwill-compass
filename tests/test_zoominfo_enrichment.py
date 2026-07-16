from types import SimpleNamespace

from darwill_ai_prospector.services.zoominfo_enrichment import (
    apply_enrichment_record,
    best_enrichment_record,
    classify_enrichment_failure,
    recursive_email_value,
)


def test_gabe_style_response_is_parsed():
    payload = {
        "contact_1": {
            "success": True,
            "data": {
                "firstName": "Gabe",
                "lastName": "Garwick",
                "companyName": "Mr. Electric",
                "jobTitle": "Marketing Specialist",
                "email": "gabriel@mrelectricok.com",
                "mobilePhone": "(918) 693-1561",
                "matchStatus": "FULL_MATCH",
                "zoominfoCompanyId": 25756514,
            },
        }
    }
    contact = SimpleNamespace(
        first_name="Gabe",
        last_name="Garwick",
        company_name="Mr. Electric",
        title="Marketing Specialist",
        email="",
        direct_phone="",
        mobile_phone="",
        email_status="",
        email_confidence=0,
        email_verification_status="",
        email_recovery_method="",
        phone_status="",
        phone_confidence=0,
        zoominfo_enriched_company_id="",
        zoominfo_match_status="",
        zoominfo_retry_message="",
        zoominfo_enrichment_warnings="",
        zoominfo_enrichment_result="",
    )

    record = best_enrichment_record(payload, contact)
    apply_enrichment_record(contact, record)

    assert contact.email == "gabriel@mrelectricok.com"
    assert contact.mobile_phone == "(918) 693-1561"
    assert contact.zoominfo_match_status == "FULL_MATCH"
    assert contact.email_confidence == 99


def test_limit_exceeded_is_classified():
    payload = {
        "contact_1": {
            "success": False,
            "error": "Contact enrichment failed: Limit exceeded",
        }
    }
    assert classify_enrichment_failure(
        payload,
        person_id="",
        extracted_email="",
    ) == "ZoomInfo MCP enrichment limit exceeded"


def test_recursive_email_value():
    email, path = recursive_email_value(
        {"contact_1": {"data": {"businessEmail": "a@example.com"}}}
    )
    assert email == "a@example.com"
    assert path.endswith("businessEmail")
