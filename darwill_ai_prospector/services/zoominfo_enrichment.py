"""ZoomInfo contact-enrichment parsing and diagnostics.

This module contains no Tkinter or application state. It can be tested
independently from the desktop UI.
"""

from __future__ import annotations

import json
import re
from typing import Any


EMAIL_FIELD_NAMES = {
    "email",
    "emailaddress",
    "businessemail",
    "workemail",
    "primaryemail",
    "verifiedemail",
    "professionalemail",
    "contactemail",
    "personemail",
}


def _text_blob(value: Any) -> str:
    try:
        return json.dumps(value, default=str)
    except Exception:
        return str(value)


def collect_email_diagnostics(
    value: Any,
    path: str = "",
) -> list[dict[str, str]]:
    """Return all email-related fields found in a nested MCP payload."""
    found: list[dict[str, str]] = []
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}" if path else str(key)
            normalized = re.sub(r"[^a-z0-9]", "", str(key).lower())
            if normalized in EMAIL_FIELD_NAMES:
                found.append(
                    {
                        "path": child_path,
                        "field": str(key),
                        "value": str(child or "").strip(),
                    }
                )
            found.extend(collect_email_diagnostics(child, child_path))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            found.extend(
                collect_email_diagnostics(child, f"{path}[{index}]")
            )
    return found


def recursive_email_value(value: Any) -> tuple[str, str]:
    """Extract the first syntactically valid email from a nested payload."""
    for candidate in collect_email_diagnostics(value):
        email = candidate["value"].strip()
        if re.fullmatch(
            r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
            email,
            re.I,
        ):
            return email, candidate["path"]

    blob = _text_blob(value)
    match = re.search(
        r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
        blob,
        re.I,
    )
    return (match.group(0), "payload_text_scan") if match else ("", "")


def normalize_contact_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def enrichment_result_records(payload: Any) -> list[dict[str, Any]]:
    """Extract successful `contact_*.data` records from MCP output."""
    records: list[dict[str, Any]] = []
    if not isinstance(payload, dict):
        return records

    for value in payload.values():
        if not isinstance(value, dict):
            continue
        data = value.get("data")
        if value.get("success") is True and isinstance(data, dict):
            records.append(data)
        elif isinstance(data, dict) and data.get("success") is True:
            records.append(data)
    return records


def best_enrichment_record(
    payload: Any,
    contact: Any,
) -> dict[str, Any]:
    """Choose the response record most likely to represent `contact`."""
    records = enrichment_result_records(payload)
    if not records:
        return {}

    target_first = normalize_contact_name(contact.first_name)
    target_last = normalize_contact_name(contact.last_name)
    target_company = normalize_contact_name(contact.company_name)
    target_title = normalize_contact_name(contact.title)

    def score(record: dict[str, Any]) -> int:
        points = 0
        if normalize_contact_name(record.get("firstName", "")) == target_first:
            points += 4
        if normalize_contact_name(record.get("lastName", "")) == target_last:
            points += 4
        if normalize_contact_name(record.get("companyName", "")) == target_company:
            points += 3
        if normalize_contact_name(
            record.get("jobTitle", record.get("title", ""))
        ) == target_title:
            points += 2
        if str(record.get("matchStatus", "")).upper() == "FULL_MATCH":
            points += 5
        if record.get("email"):
            points += 5
        return points

    return max(records, key=score)


def apply_enrichment_record(
    contact: Any,
    record: dict[str, Any],
) -> None:
    """Apply a successful ZoomInfo enrichment record to a contact model."""
    if not record:
        return

    email = str(record.get("email", "") or "").strip()
    mobile = str(record.get("mobilePhone", "") or "").strip()
    phone = str(record.get("phone", "") or "").strip()
    company_id = str(
        record.get("zoominfoCompanyId", record.get("companyId", "")) or ""
    ).strip()
    match_status = str(record.get("matchStatus", "") or "").strip()
    retry_info = record.get("retryInfo", {})
    warnings = record.get("warnings", [])

    if email:
        contact.email = email
        contact.email_status = "ZoomInfo enrichment returned"
        contact.email_confidence = 99 if match_status == "FULL_MATCH" else 96
        contact.email_verification_status = "ZoomInfo Enriched"
        contact.email_recovery_method = (
            "ZoomInfo identity-field enrichment "
            f"({match_status or 'matched'})"
        )

    if mobile:
        contact.mobile_phone = mobile
        contact.phone_status = "ZoomInfo enrichment returned"
        contact.phone_confidence = 96
    elif phone:
        contact.direct_phone = phone
        contact.phone_status = "ZoomInfo enrichment returned"
        contact.phone_confidence = 94

    if company_id:
        contact.zoominfo_enriched_company_id = company_id

    contact.zoominfo_match_status = match_status
    contact.zoominfo_retry_message = str(
        retry_info.get("message", "") if isinstance(retry_info, dict) else ""
    )
    contact.zoominfo_enrichment_warnings = (
        json.dumps(warnings, default=str) if warnings else ""
    )
    contact.zoominfo_enrichment_result = (
        "Verified email recovered"
        if email
        else "Contact enriched but no email returned"
    )


def classify_enrichment_failure(
    payload: Any,
    *,
    person_id: str,
    extracted_email: str,
) -> str:
    """Return a user-facing classification for an unsuccessful response."""
    if extracted_email:
        return "Email recovered"

    blob = _text_blob(payload).lower()
    if "limit exceeded" in blob:
        return "ZoomInfo MCP enrichment limit exceeded"
    if any(
        term in blob
        for term in [
            "not entitled",
            "entitlement",
            "permission denied",
            "forbidden",
            "unauthorized",
        ]
    ):
        return "ZoomInfo entitlement or permission issue"
    if any(
        term in blob
        for term in [
            "rate limit",
            "too many requests",
            "throttl",
        ]
    ):
        return "ZoomInfo rate limit"
    if any(
        term in blob
        for term in [
            "invalid person id",
            "invalid personid",
            "invalid contact id",
            "contact not found",
            "person not found",
            "no matching contact",
            "unable to match contact",
        ]
    ):
        return f"Invalid or unmatched ZoomInfo person ID: {person_id}"
    if not enrichment_result_records(payload):
        return "Enrichment returned no successful contact records"
    return "Enrichment returned a contact record but no email field"
