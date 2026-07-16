#!/usr/bin/env python3
from __future__ import annotations

import asyncio
import base64
import hashlib
import json
import queue
import re
import secrets
import sqlite3
import threading
import time
import traceback
import webbrowser
import os
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from datetime import datetime, timezone, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urlparse, parse_qs, urljoin
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import httpx
import keyring
import requests
import csv
import shutil
from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client
from tavily import TavilyClient
import xlsxwriter
from bs4 import BeautifulSoup
from pypdf import PdfReader
import dns.resolver
from cryptography.fernet import Fernet, InvalidToken

from .ui.scrolling import (
    VerticalScrolledFrame,
    configure_scrollable_text,
    configure_scrollable_tree,
)
from .ui.company_intelligence import (
    build_company_intelligence as build_company_presentation,
)
from .ui.theme import (
    apply_compass_theme,
    workspace_palette,
)
from .ui.mcp_explorer import open_mcp_explorer_window
from .services.zoominfo_enrichment import (
    apply_enrichment_record,
    best_enrichment_record,
    classify_enrichment_failure,
    collect_email_diagnostics,
    enrichment_result_records,
    recursive_email_value,
)
from .services.zoominfo_adapter import (
    ZoomInfoAdapter,
    ZoomInfoAdapterError,
)
from .services.company_intelligence import (
    build_company_intelligence as build_company_brief,
    build_email_resolution,
    format_company_intelligence,
)
from .services.intelligence_store import IntelligenceStore
from .services.email_intelligence import (
    build_email_intelligence,
    format_email_intelligence_report,
)

APP_TITLE = "Darwill Compass 8.6 — Expandable Tab Workspace"
SERVICE = "DarwillProspectIntelligence"
BASE_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = BASE_DIR / "settings.json"
PROFILES_FILE = BASE_DIR / "search_profiles.json"
TRADE_PRESETS_FILE = BASE_DIR / "trade_presets.json"
LEARNING_FILE = BASE_DIR / "qualification_learning.json"
DARWILL_KNOWLEDGE_FILE = BASE_DIR / "darwill_knowledge.json"
OUTREACH_QUEUE_FILE = BASE_DIR / "outreach_queue.json"
DELIVERABILITY_RULES_FILE = BASE_DIR / "deliverability_rules.json"
HUBSPOT_EXPORT_DIR = BASE_DIR.parent / "hubspot_exports"
HUBSPOT_EXPORT_DIR.mkdir(exist_ok=True)
MASTER_CSV_BACKUP_DIR = BASE_DIR.parent / "master_csv_backups"
MASTER_CSV_BACKUP_DIR.mkdir(exist_ok=True)
DEFAULT_MASTER_CSV_NAME = "Darwill_Master_Prospect_Database.csv"
CHECKPOINT_FILE = BASE_DIR / "run_checkpoint.json"
OUTPUT_DIR = BASE_DIR.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
DB_PATH = BASE_DIR / "prospector.db"
LOG_DIR = BASE_DIR / "logs"
TOKEN_FILE = BASE_DIR / "oauth_tokens.enc"
LOG_DIR.mkdir(exist_ok=True)

MCP_URL = "https://mcp.zoominfo.com/mcp"
AUTHORIZE_URL = "https://api.zoominfo.com/gtm/oauth/v1/authorize"
TOKEN_URL = "https://api.zoominfo.com/gtm/oauth/v1/token"
REDIRECT_URI = "http://localhost:8080/callback"

# Darwill-inspired blue palette
BLUE_DARK = "#003B71"
BLUE = "#005EB8"
BLUE_MID = "#1874C9"
BLUE_LIGHT = "#E8F2FC"
BLUE_PALE = "#F5F9FD"
WHITE = "#FFFFFF"
TEXT = "#17324D"
MUTED = "#5A7184"
SUCCESS = "#16794C"
WARNING = "#A05A00"
ERROR = "#B42318"
NAVY = "#082B4C"
NAVY_2 = "#0E3A63"
ACCENT = "#0B6FD3"
ACCENT_HOVER = "#0A5EAF"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#F4F8FC"
BORDER = "#D8E4EF"
SHADOW = "#C9D8E6"
SIDEBAR = "#071B2E"
SIDEBAR_SECTION = "#0A243D"
SIDEBAR_HOVER = "#123A5F"
SIDEBAR_ACTIVE = "#1F6FD1"
CONTENT_BG = "#EEF3F8"
PRODUCT_VERSION = "8.6"
DEVELOPER_NAME = "Jon Tigchelaar"

CONTACT_SOURCE_PRIORITY = {
    "ZoomInfo": 100,
    "Company Website": 92,
    "Press Release": 84,
    "Public Profile": 78,
    "Public Search": 70,
    "Predicted": 45,
}

PUBLIC_CONTACT_PATTERNS = [
    r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
    r"\b(?:\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}\b",
]

DECISION_MAKER_TITLE_TERMS = [
    "chief marketing officer", "cmo", "vice president of marketing",
    "vp marketing", "vice president marketing", "director of marketing",
    "marketing director", "head of marketing", "growth marketing",
    "director of growth", "vice president of growth", "vp growth",
    "chief growth officer", "chief revenue officer", "cro",
    "president", "chief executive officer", "ceo", "owner",
    "chief operating officer", "coo", "general manager",
    "business development", "demand generation", "brand",
]

LOW_VALUE_PUBLIC_EMAIL_PREFIXES = {
    "info", "contact", "support", "service", "help", "hello",
    "office", "admin", "careers", "jobs", "sales",
}


def public_contact_type(value: str) -> str:
    value = (value or "").strip()
    if "@" in value:
        return "email"
    if re.search(r"\d{3}", value):
        return "phone"
    return "other"


def normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) != 10:
        return value.strip()
    return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"


def extract_public_contacts(text: str) -> tuple[list[str], list[str]]:
    emails = sorted(set(
        match.lower()
        for match in re.findall(
            PUBLIC_CONTACT_PATTERNS[0], text or "", flags=re.I
        )
    ))
    phones = sorted(set(
        normalize_phone(match)
        for match in re.findall(
            PUBLIC_CONTACT_PATTERNS[1], text or "", flags=re.I
        )
    ))
    return emails, phones


def email_local_part(email: str) -> str:
    return email.split("@", 1)[0].strip().lower() if "@" in email else ""


def company_domain_from_prospect(prospect: "Prospect") -> str:
    return normalize_domain(prospect.website)


def title_relevance_score(title: str) -> int:
    lowered = (title or "").lower()
    weights = [
        ("chief marketing officer", 150), ("cmo", 150),
        ("vice president of marketing", 145), ("vp marketing", 145),
        ("director of marketing", 138), ("marketing director", 138),
        ("head of marketing", 136), ("vice president of growth", 136),
        ("director of growth", 134), ("chief growth officer", 134),
        ("chief revenue officer", 128), ("demand generation", 126),
        ("growth marketing", 124), ("brand director", 120),
        ("marketing manager", 116), ("business development", 108),
        ("president", 96), ("chief executive officer", 94),
        ("ceo", 94), ("owner", 92), ("chief operating officer", 80),
        ("coo", 80), ("general manager", 65),
    ]
    for phrase, score in weights:
        if phrase in lowered:
            return score
    return 0


def email_pattern_from_known_contacts(
    contacts: list["RankedContact"],
    domain: str,
) -> tuple[str, float]:
    examples = []
    for contact in contacts:
        email = (contact.email or "").strip().lower()
        if not email or normalize_domain(email.split("@")[-1]) != domain:
            continue
        local = email_local_part(email)
        first = re.sub(r"[^a-z]", "", contact.first_name.lower())
        last = re.sub(r"[^a-z]", "", contact.last_name.lower())
        if not first or not last:
            continue
        patterns = {
            "first.last": f"{first}.{last}",
            "firstlast": f"{first}{last}",
            "flast": f"{first[:1]}{last}",
            "firstl": f"{first}{last[:1]}",
            "last.first": f"{last}.{first}",
            "lastf": f"{last}{first[:1]}",
            "first": first,
            "last": last,
        }
        for name, expected in patterns.items():
            if local == expected:
                examples.append(name)
                break
    if not examples:
        return "", 0.0
    counts = {}
    for name in examples:
        counts[name] = counts.get(name, 0) + 1
    pattern, count = max(counts.items(), key=lambda item: item[1])
    confidence = min(0.95, 0.55 + 0.15 * count)
    return pattern, confidence


def predict_email(
    first_name: str,
    last_name: str,
    domain: str,
    pattern: str,
) -> str:
    first = re.sub(r"[^a-z]", "", (first_name or "").lower())
    last = re.sub(r"[^a-z]", "", (last_name or "").lower())
    if not first or not last or not domain or not pattern:
        return ""
    local_patterns = {
        "first.last": f"{first}.{last}",
        "firstlast": f"{first}{last}",
        "flast": f"{first[:1]}{last}",
        "firstl": f"{first}{last[:1]}",
        "last.first": f"{last}.{first}",
        "lastf": f"{last}{first[:1]}",
        "first": first,
        "last": last,
    }
    local = local_patterns.get(pattern, "")
    return f"{local}@{domain}" if local else ""


def unique_public_contact_id(
    company_id: str,
    full_name: str,
    title: str,
) -> str:
    payload = f"{company_id}|{full_name.lower()}|{title.lower()}".encode("utf-8")
    return "public-" + hashlib.sha1(payload).hexdigest()[:16]


PROSPECT_STATUSES = [
    "New",
    "Researching",
    "Qualified",
    "Rejected",
    "Approved",
    "Synced to HubSpot",
    "In Sequence",
    "Meeting",
    "Customer",
    "Lost",
]

TERMINAL_SKIP_STATUSES = {
    "Qualified",
    "Approved",
    "Synced to HubSpot",
    "In Sequence",
    "Meeting",
    "Customer",
    "Lost",
}

DEFAULT_STATES = "TX,AZ,CO,OK,NM,NV,UT,ID,WY,MT"
DEFAULT_TRADES = "HVAC, plumbing, electrical, pest control, pool service, garage door"

DEFAULT_TRADE_PRESETS = {
    "HVAC": {
        "keywords": "HVAC, heating, air conditioning, furnace, heat pump",
        "naics": ["238220"],
    },
    "Plumbing": {
        "keywords": "plumbing, plumber, drain, sewer, water heater, repiping",
        "naics": ["238220"],
    },
    "Electrical": {
        "keywords": "electrical contractor, electrician, electrical repair, generator installation",
        "naics": ["238210"],
    },
    "Pest Control": {
        "keywords": "pest control, exterminator, termite, rodent control",
        "naics": ["561710"],
    },
    "Pool Service": {
        "keywords": "pool service, swimming pool maintenance, pool repair, pool cleaning",
        "naics": ["561790"],
    },
    "Garage Door": {
        "keywords": "garage door repair, garage door installation, overhead door service",
        "naics": ["238290"],
    },
    "Solar": {
        "keywords": "residential solar, solar installation, rooftop solar",
        "naics": ["238210"],
    },
    "Remodeling": {
        "keywords": "residential remodeling, home renovation, kitchen remodeling, bathroom remodeling, whole-home remodeling",
        "naics": ["236118"],
    },
    "Landscaping": {
        "keywords": "residential landscaping, lawn care, landscape maintenance",
        "naics": ["561730"],
    },
}

TRADE_SUGGESTION_LIBRARY = {
    "remodeling": {
        "keywords": "residential remodeling, home renovation, kitchen remodeling, bathroom remodeling, whole-home remodeling",
        "naics": ["236118"],
        "confidence": 98,
        "note": "Residential Remodelers is the primary classification.",
    },
    "roofing": {
        "keywords": "residential roofing, roof replacement, roof repair, re-roofing",
        "naics": ["238160"],
        "confidence": 98,
        "note": "Roofing Contractors is a direct classification.",
    },
    "window replacement": {
        "keywords": "replacement windows, residential windows, window installation, entry doors",
        "naics": ["238350"],
        "confidence": 82,
        "note": "Finish Carpentry Contractors is commonly used; review mixed product/install businesses.",
    },
    "windows": {
        "keywords": "replacement windows, residential windows, window installation",
        "naics": ["238350"],
        "confidence": 78,
        "note": "Retailers and manufacturers must still be excluded.",
    },
    "water treatment": {
        "keywords": "residential water treatment, water softener installation, water filtration, reverse osmosis",
        "naics": ["238220"],
        "confidence": 72,
        "note": "Installation companies often sit under plumbing; manufacturers should be excluded.",
    },
    "foundation repair": {
        "keywords": "foundation repair, residential foundation repair, basement waterproofing, crawl space repair",
        "naics": ["238110", "238990"],
        "confidence": 74,
        "note": "Multiple specialty-contractor classifications are common.",
    },
    "restoration": {
        "keywords": "water damage restoration, fire damage restoration, mold remediation, emergency restoration",
        "naics": ["562910"],
        "confidence": 90,
        "note": "Remediation Services is the standard classification.",
    },
    "painting": {
        "keywords": "residential painting, interior painting, exterior painting, house painting",
        "naics": ["238320"],
        "confidence": 98,
        "note": "Painting and Wall Covering Contractors is a direct classification.",
    },
    "flooring": {
        "keywords": "residential flooring, hardwood installation, carpet installation, tile flooring",
        "naics": ["238330"],
        "confidence": 95,
        "note": "Flooring Contractors is a direct classification.",
    },
    "insulation": {
        "keywords": "residential insulation, attic insulation, spray foam insulation, home insulation",
        "naics": ["238310"],
        "confidence": 96,
        "note": "Drywall and Insulation Contractors is the standard classification.",
    },
    "appliance repair": {
        "keywords": "appliance repair, refrigerator repair, washer repair, residential appliance service",
        "naics": ["811412"],
        "confidence": 98,
        "note": "Appliance Repair and Maintenance is a direct classification.",
    },
    "chimney": {
        "keywords": "chimney sweep, chimney repair, fireplace service, chimney inspection",
        "naics": ["561790"],
        "confidence": 66,
        "note": "Often grouped under Other Services to Buildings; review manually.",
    },
    "fencing": {
        "keywords": "residential fencing, fence installation, fence repair, privacy fence",
        "naics": ["238990"],
        "confidence": 72,
        "note": "Usually classified under Other Specialty Trade Contractors.",
    },
    "tree service": {
        "keywords": "tree service, tree removal, arborist, stump grinding, residential tree care",
        "naics": ["561730"],
        "confidence": 92,
        "note": "Landscaping Services commonly covers tree-care operators.",
    },
    "house cleaning": {
        "keywords": "house cleaning, residential cleaning, maid service, home cleaning",
        "naics": ["561720"],
        "confidence": 98,
        "note": "Janitorial Services is the standard classification.",
    },
    "cleaning": {
        "keywords": "house cleaning, residential cleaning, maid service",
        "naics": ["561720"],
        "confidence": 88,
        "note": "Confirm the company is residential rather than commercial janitorial.",
    },
}


DEFAULT_DELIVERABILITY_RULES = {
    "minimum_enrollment_score": 82,
    "maximum_words": 145,
    "preferred_words": 105,
    "maximum_links": 1,
    "maximum_exclamation_marks": 0,
    "require_company_name": True,
    "require_contact_first_name": True,
    "require_single_clear_cta": True,
    "sender_authentication": {
        "spf_confirmed": True,
        "dkim_confirmed": True,
        "dmarc_confirmed": True,
    },
    "risky_phrases": [
        "free", "guaranteed", "limited time", "act now", "click here",
        "urgent", "buy now", "special promotion", "no obligation",
        "risk free", "winner", "cash", "lowest price", "100%",
        "exclusive offer", "once in a lifetime",
    ],
    "generic_openers": [
        "i hope this email finds you well",
        "i wanted to reach out",
        "i'm reaching out",
        "i came across your company",
        "i was impressed by your company",
        "quick question",
    ],
    "salesy_phrases": [
        "revolutionize", "game-changing", "best-in-class", "cutting-edge",
        "unlock your potential", "supercharge", "skyrocket",
        "transform your business", "industry-leading",
    ],
}


def load_deliverability_rules() -> dict[str, Any]:
    if DELIVERABILITY_RULES_FILE.exists():
        try:
            data = json.loads(
                DELIVERABILITY_RULES_FILE.read_text(encoding="utf-8")
            )
            if isinstance(data, dict):
                merged = json.loads(json.dumps(DEFAULT_DELIVERABILITY_RULES))
                for key, value in data.items():
                    if key == "sender_authentication" and isinstance(value, dict):
                        merged[key].update(value)
                    else:
                        merged[key] = value
                return merged
        except Exception:
            pass
    DELIVERABILITY_RULES_FILE.write_text(
        json.dumps(DEFAULT_DELIVERABILITY_RULES, indent=2),
        encoding="utf-8",
    )
    return json.loads(json.dumps(DEFAULT_DELIVERABILITY_RULES))


def _count_links(text: str) -> int:
    return len(re.findall(r"https?://|www\.", text or "", flags=re.I))


def _sentence_count(text: str) -> int:
    pieces = re.split(r"[.!?]+(?:\s+|$)", (text or "").strip())
    return max(1, len([piece for piece in pieces if piece.strip()]))


def _word_count(text: str) -> int:
    return len(re.findall(r"\b[\w'-]+\b", text or ""))


def _uppercase_ratio(text: str) -> float:
    letters = [char for char in (text or "") if char.isalpha()]
    if not letters:
        return 0.0
    return sum(char.isupper() for char in letters) / len(letters)


def _cta_count(text: str) -> int:
    patterns = [
        r"\bwould you be open\b",
        r"\bdo you have\b.{0,35}\bminutes\b",
        r"\bcan we\b.{0,35}\b(?:talk|connect|meet)\b",
        r"\bschedule\b",
        r"\bbook\b.{0,20}\btime\b",
        r"\bgrab a time\b",
        r"\blet me know\b",
        r"\binterested\b",
    ]
    lowered = (text or "").lower()
    return sum(bool(re.search(pattern, lowered)) for pattern in patterns)


def safer_subjects(company_name: str, strategy: str) -> list[str]:
    company = company_name.strip() or "your company"
    strategy_low = (strategy or "").lower()
    if "new mover" in strategy_low:
        candidates = [
            f"New homeowner opportunity for {company}",
            f"Territory insight for {company}",
            f"Report I created for {company}",
        ]
    elif "data" in strategy_low or "attribution" in strategy_low:
        candidates = [
            f"Audience insight for {company}",
            f"Measuring customer acquisition at {company}",
            f"Data idea for {company}",
        ]
    elif "expansion" in strategy_low or "brand" in strategy_low:
        candidates = [
            f"Supporting {company}'s market growth",
            f"An idea for {company}'s service areas",
            f"Growth opportunity for {company}",
        ]
    elif "retention" in strategy_low:
        candidates = [
            f"Customer reactivation idea for {company}",
            f"Growing repeat business at {company}",
            f"Retention opportunity for {company}",
        ]
    else:
        candidates = [
            f"An idea for {company}",
            f"Thought this might be useful for {company}",
            f"Marketing opportunity for {company}",
        ]
    return candidates


def analyze_deliverability(
    subject: str,
    body: str,
    company_name: str,
    contact_first_name: str,
    strategy: str,
    rules: dict[str, Any],
) -> dict[str, Any]:
    subject = (subject or "").strip()
    body = (body or "").strip()
    lowered_subject = subject.lower()
    lowered_body = body.lower()
    combined = f"{subject}\n{body}".lower()

    issues: list[str] = []
    recommendations: list[str] = []
    positives: list[str] = []

    subject_score = 100
    personalization_score = 100
    content_score = 100
    human_tone_score = 100
    format_score = 100

    words = _word_count(body)
    links = _count_links(body)
    exclamations = body.count("!") + subject.count("!")
    ctas = _cta_count(body)
    upper_ratio = _uppercase_ratio(subject + " " + body)

    risky_hits = [
        phrase for phrase in rules.get("risky_phrases", [])
        if phrase.lower() in combined
    ]
    generic_hits = [
        phrase for phrase in rules.get("generic_openers", [])
        if phrase.lower() in lowered_body
    ]
    salesy_hits = [
        phrase for phrase in rules.get("salesy_phrases", [])
        if phrase.lower() in combined
    ]

    # Subject quality
    if not subject:
        subject_score -= 60
        issues.append("Subject line is blank.")
        recommendations.append("Use a short, specific subject tied to the company.")
    if len(subject) > 58:
        subject_score -= min(25, (len(subject) - 58) // 2 + 8)
        issues.append(f"Subject is long ({len(subject)} characters).")
        recommendations.append("Keep the subject near 35–55 characters.")
    elif 20 <= len(subject) <= 55:
        positives.append("Subject length is concise.")
    if exclamations:
        subject_score -= min(20, exclamations * 8)
        issues.append("Subject/body uses exclamation marks.")
        recommendations.append("Remove exclamation marks from cold outreach.")
    if re.search(r"\b(?:free|urgent|guaranteed|act now|limited time)\b", lowered_subject):
        subject_score -= 25
        issues.append("Subject contains promotional language.")
    if subject.isupper() and subject:
        subject_score -= 35
        issues.append("Subject is written in all caps.")
    if company_name and company_name.lower() in lowered_subject:
        subject_score += 3
        positives.append("Subject is company-specific.")

    # Personalization
    if rules.get("require_contact_first_name", True):
        if contact_first_name and re.search(
            rf"\b{re.escape(contact_first_name.lower())}\b", lowered_body
        ):
            positives.append("Contact's first name appears in the email.")
        else:
            personalization_score -= 22
            issues.append("Contact's first name is not used.")
            recommendations.append("Address the contact by first name.")
    if rules.get("require_company_name", True):
        if company_name and company_name.lower() in lowered_body:
            positives.append("Company name appears in the email.")
        else:
            personalization_score -= 25
            issues.append("Company name is not used in the body.")
            recommendations.append("Reference the company naturally in the opening.")
    evidence_terms = [
        "noticed", "researching", "expanded", "location", "service area",
        "marketing", "growth", "membership", "technology", "territory",
        "homeowner", "customer acquisition",
    ]
    evidence_count = sum(term in lowered_body for term in evidence_terms)
    if evidence_count >= 2:
        positives.append("Opening contains specific company or market context.")
    elif evidence_count == 0:
        personalization_score -= 18
        issues.append("The email lacks a clear researched detail.")
        recommendations.append(
            "Use one verified company signal rather than a generic compliment."
        )

    # Content and spam-risk heuristics
    if risky_hits:
        content_score -= min(40, len(risky_hits) * 10)
        issues.append("Risky promotional phrases: " + ", ".join(risky_hits))
        recommendations.append("Replace promotional claims with factual language.")
    if salesy_hits:
        human_tone_score -= min(35, len(salesy_hits) * 9)
        issues.append("Overly promotional phrasing: " + ", ".join(salesy_hits))
        recommendations.append("Use plain, concrete language.")
    if generic_hits:
        human_tone_score -= min(30, len(generic_hits) * 12)
        issues.append("Generic cold-email opener detected.")
        recommendations.append("Lead with a verified company-specific observation.")
    if upper_ratio > 0.18:
        content_score -= 18
        issues.append("Unusually high capitalization.")
    if words > int(rules.get("maximum_words", 145)):
        deduction = min(30, math.ceil((words - int(rules["maximum_words"])) / 8))
        content_score -= deduction
        issues.append(f"Email is long ({words} words).")
        recommendations.append(
            f"Reduce the email to roughly {rules.get('preferred_words', 105)} words."
        )
    elif 55 <= words <= int(rules.get("maximum_words", 145)):
        positives.append(f"Email length is reasonable ({words} words).")
    elif words < 35:
        content_score -= 10
        issues.append("Email may be too brief to establish relevance.")

    # Formatting, links and CTA
    max_links = int(rules.get("maximum_links", 1))
    if links > max_links:
        format_score -= min(35, (links - max_links) * 15)
        issues.append(f"Email contains {links} links.")
        recommendations.append("Use zero links in the first email, or at most one.")
    elif links == 0:
        positives.append("No external links in the body.")
    if "<img" in lowered_body or "data:image" in lowered_body:
        format_score -= 30
        issues.append("Image or embedded-image markup detected.")
        recommendations.append("Use a plain-text first touch without images.")
    if "<table" in lowered_body or "<div" in lowered_body or "<span" in lowered_body:
        format_score -= 18
        issues.append("HTML-heavy formatting detected.")
        recommendations.append("Use simple paragraphs and a plain signature.")
    if ctas == 1:
        positives.append("Email contains one clear CTA.")
    elif ctas == 0:
        content_score -= 12
        issues.append("No clear next step was detected.")
        recommendations.append("End with one low-friction question.")
    else:
        content_score -= min(22, (ctas - 1) * 10)
        issues.append("Multiple calls to action may feel automated.")
        recommendations.append("Keep one CTA: a brief 15-minute conversation.")

    sentence_count = _sentence_count(body)
    avg_sentence_words = words / max(1, sentence_count)
    if avg_sentence_words > 24:
        human_tone_score -= 15
        issues.append("Sentences are dense.")
        recommendations.append("Shorten long sentences and vary sentence length.")
    elif 8 <= avg_sentence_words <= 20:
        positives.append("Sentence length is conversational.")

    # Authentication is a manual confirmation; the app cannot infer live DNS here.
    auth = rules.get("sender_authentication", {})
    auth_count = sum(bool(auth.get(key)) for key in [
        "spf_confirmed", "dkim_confirmed", "dmarc_confirmed"
    ])
    authentication_score = round(auth_count / 3 * 100)
    if auth_count < 3:
        issues.append("Sender authentication checklist is incomplete.")
        recommendations.append("Confirm SPF, DKIM, and DMARC before sending.")

    subject_score = max(0, min(100, subject_score))
    personalization_score = max(0, min(100, personalization_score))
    content_score = max(0, min(100, content_score))
    human_tone_score = max(0, min(100, human_tone_score))
    format_score = max(0, min(100, format_score))

    inbox_score = round(
        subject_score * 0.18
        + personalization_score * 0.23
        + content_score * 0.20
        + human_tone_score * 0.15
        + format_score * 0.14
        + authentication_score * 0.10
    )
    risk = "Low" if inbox_score >= 88 else "Moderate" if inbox_score >= 75 else "High"
    ready = inbox_score >= int(rules.get("minimum_enrollment_score", 82))

    return {
        "inbox_readiness_score": inbox_score,
        "subject_quality_score": subject_score,
        "personalization_score": personalization_score,
        "content_risk_score": content_score,
        "human_tone_score": human_tone_score,
        "formatting_score": format_score,
        "authentication_score": authentication_score,
        "spam_risk": risk,
        "send_ready": ready,
        "word_count": words,
        "link_count": links,
        "cta_count": ctas,
        "issues": issues,
        "recommendations": list(dict.fromkeys(recommendations)),
        "positives": list(dict.fromkeys(positives)),
        "subject_alternatives": safer_subjects(company_name, strategy),
    }


def safer_email_rewrite(
    body: str,
    company_name: str,
    contact_first_name: str,
    strategy: str,
) -> str:
    lines = [line.strip() for line in (body or "").splitlines() if line.strip()]
    first = contact_first_name.strip() or "there"
    company = company_name.strip() or "your company"

    researched = ""
    for line in lines:
        lower = line.lower()
        if company.lower() in lower and any(
            term in lower for term in [
                "noticed", "research", "expanded", "growth", "service area",
                "marketing", "location", "homeowner", "territory",
            ]
        ):
            researched = line
            break

    if not researched:
        researched = (
            f"I was researching {company} and thought the work your team is doing "
            "in the residential market made this worth sharing."
        )

    strategy_low = (strategy or "").lower()
    if "new mover" in strategy_low:
        value = (
            "Darwill helps home-service companies reach households shortly after "
            "they move, when they are choosing new providers."
        )
    elif "data" in strategy_low or "attribution" in strategy_low:
        value = (
            "Darwill combines audience data, direct mail, digital execution, and "
            "measurement to make customer acquisition easier to evaluate."
        )
    elif "retention" in strategy_low:
        value = (
            "Darwill supports customer reactivation, retention, and cross-sell "
            "programs using coordinated data-driven outreach."
        )
    elif "expansion" in strategy_low or "brand" in strategy_low:
        value = (
            "Darwill helps growing companies build awareness in precise service "
            "areas through coordinated direct mail, digital media, and measurement."
        )
    else:
        value = (
            "Darwill supports customer acquisition across audience data, direct "
            "mail, digital media, campaign execution, and attribution."
        )

    return (
        f"Hi {first},\n\n"
        f"{researched}\n\n"
        f"{value}\n\n"
        "Would you be open to a brief 15-minute conversation to see whether this "
        f"could be useful for {company}?\n\n"
        "Thank you,\nJon"
    )


DEFAULT_DARWILL_KNOWLEDGE = {
    "company_summary": (
        "Darwill is a data-driven marketing partner that helps businesses acquire, "
        "retain, reactivate, and grow customer relationships through coordinated "
        "direct mail, digital advertising, audience intelligence, analytics, "
        "campaign execution, and measurable attribution."
    ),
    "meeting_cta": (
        "Would you be open to a quick 15-minute conversation so I can show you "
        "the opportunity and explain how Darwill could support your growth goals?"
    ),
    "offerings": {
        "New Movers": {
            "value_proposition": (
                "Reach households shortly after they move, when they are actively "
                "choosing new home-service providers and establishing long-term relationships."
            ),
            "proof_points": (
                "Automated direct mail and digital touches can be coordinated by territory, "
                "with audience counts and measurable campaign reporting."
            ),
        },
        "Full-Service Marketing Partner": {
            "value_proposition": (
                "Support customer acquisition across direct mail, digital media, "
                "audience targeting, creative production, execution, and measurement."
            ),
            "proof_points": (
                "Darwill can complement an existing marketing team and fill gaps between "
                "strategy, data, production, deployment, and attribution."
            ),
        },
        "Data & Attribution": {
            "value_proposition": (
                "Use household and customer data to define audiences, improve targeting, "
                "connect campaigns to outcomes, and make marketing performance easier to measure."
            ),
            "proof_points": (
                "Darwill can support audience modeling, CRM-informed targeting, response tracking, "
                "and closed-loop campaign measurement."
            ),
        },
        "Market Expansion & Brand Awareness": {
            "value_proposition": (
                "Build awareness quickly in new service areas or recently opened markets using "
                "coordinated direct mail, digital advertising, and geographically precise audiences."
            ),
            "proof_points": (
                "Territory-level targeting helps growing operators introduce the brand to the "
                "right households while supporting branch and service-area expansion."
            ),
        },
        "Retention & Reactivation": {
            "value_proposition": (
                "Increase repeat business by reactivating past customers, supporting membership "
                "programs, and delivering timely cross-sell and lifecycle communications."
            ),
            "proof_points": (
                "Customer-file segmentation and coordinated outreach can support maintenance plans, "
                "seasonal reminders, service reminders, and cross-trade growth."
            ),
        },
    },
}

STRATEGY_NAMES = [
    "New Movers",
    "Full-Service Marketing Partner",
    "Data & Attribution",
    "Market Expansion & Brand Awareness",
    "Retention & Reactivation",
]


MARKETING_MATURITY_TERMS = {
    "paid search": ["google ads", "paid search", "ppc"],
    "social advertising": ["facebook ads", "meta ads", "instagram ads"],
    "direct mail": ["direct mail", "mailer", "postcard"],
    "broadcast": ["television advertising", "tv commercial", "radio advertising"],
    "local sponsorships": ["sponsor", "community partner", "local sponsorship"],
    "online booking": ["book online", "schedule online", "online booking"],
    "membership program": ["membership plan", "maintenance plan", "service agreement"],
}

TECHNOLOGY_TERMS = {
    "ServiceTitan": ["servicetitan"],
    "Housecall Pro": ["housecall pro", "housecallpro"],
    "HubSpot": ["hubspot"],
    "Salesforce": ["salesforce"],
    "Google Analytics": ["google analytics", "gtag"],
    "CallRail": ["callrail"],
    "Podium": ["podium"],
    "Birdeye": ["birdeye"],
}

GROWTH_EVENT_TERMS = {
    "acquisition": ["acquired", "acquisition", "merged with", "merger"],
    "private equity": ["private equity", "portfolio company", "backed by"],
    "expansion": ["new location", "opened a new", "expanded into", "expansion"],
    "hiring": ["we're hiring", "now hiring", "careers"],
}
PRIMARY_TITLES = [
    ("chief marketing officer", 130), ("cmo", 130),
    ("vice president of marketing", 125), ("vp marketing", 125),
    ("vice president of growth", 122), ("vp growth", 122),
    ("director of marketing", 118), ("marketing director", 118),
    ("director of growth", 116), ("head of marketing", 116),
    ("vice president of sales and marketing", 112),
    ("vp sales and marketing", 112), ("marketing manager", 102),
    ("demand generation", 100), ("growth marketing", 100),
    ("brand director", 96), ("president", 84),
    ("chief executive officer", 84), ("ceo", 84), ("owner", 82),
    ("chief operating officer", 76), ("coo", 76), ("general manager", 72),
]
EXCLUDED_TITLES = [
    "finance", "financial", "controller", "accounting", "human resources",
    "recruiter", "legal", "attorney", "assistant", "coordinator", "intern",
    "technician", "procurement", "supply chain",
]
POSITIVE_COMPANY_TERMS = [
    "residential", "homeowner", "home services", "heating", "air conditioning",
    "hvac", "plumbing", "electrician", "electrical", "pest control",
    "pool service", "garage door", "drain", "sewer", "water heater",
]
NEGATIVE_COMPANY_TERMS = [
    "commercial only", "industrial only", "manufacturer", "manufacturing",
    "wholesale distributor", "equipment distributor", "retailer", "marketplace",
    "insurance", "bank", "hospital", "software company", "roofing", "roofers",
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_domain(value: str) -> str:
    if not value:
        return ""
    text = value.strip().lower()
    if "://" not in text:
        text = "https://" + text
    return (urlparse(text).hostname or "").removeprefix("www.").strip(".")


def text_blob(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return str(value)
    if isinstance(value, dict):
        return " ".join(f"{k} {text_blob(v)}" for k, v in value.items())
    if isinstance(value, list):
        return " ".join(text_blob(v) for v in value)
    return str(value)


def recursive_find(value: Any, aliases: set[str]) -> list[Any]:
    hits = []
    if isinstance(value, dict):
        for key, item in value.items():
            normalized = re.sub(r"[^a-z0-9]", "", key.lower())
            if normalized in aliases and item not in (None, "", [], {}):
                hits.append(item)
            hits.extend(recursive_find(item, aliases))
    elif isinstance(value, list):
        for item in value:
            hits.extend(recursive_find(item, aliases))
    return hits


def first_value(value: Any, aliases: list[str], default: Any = "") -> Any:
    normalized = {re.sub(r"[^a-z0-9]", "", a.lower()) for a in aliases}
    hits = recursive_find(value, normalized)
    for hit in hits:
        if isinstance(hit, list):
            if hit:
                return hit[0]
        elif not isinstance(hit, dict):
            return hit
    return default


STATE_NAMES_TO_CODES = {
    "texas": "TX", "arizona": "AZ", "colorado": "CO", "oklahoma": "OK",
    "new mexico": "NM", "nevada": "NV", "utah": "UT", "idaho": "ID",
    "wyoming": "WY", "montana": "MT",
    "maryland": "MD", "florida": "FL", "new york": "NY", "virginia": "VA",
    "california": "CA", "washington": "WA", "oregon": "OR",
    "pennsylvania": "PA", "ohio": "OH", "georgia": "GA",
    "north carolina": "NC", "south carolina": "SC", "tennessee": "TN",
    "missouri": "MO", "kansas": "KS", "nebraska": "NE",
}


def normalize_state(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    if len(text) == 2:
        return text.upper()
    return STATE_NAMES_TO_CODES.get(text.lower(), text.upper())


def allowed_state_codes(value: str) -> set[str]:
    return {
        normalize_state(item)
        for item in (value or "").split(",")
        if normalize_state(item)
    }


def load_trade_presets() -> dict[str, dict[str, Any]]:
    if TRADE_PRESETS_FILE.exists():
        try:
            data = json.loads(TRADE_PRESETS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data:
                cleaned = {}
                for name, preset in data.items():
                    if not isinstance(preset, dict):
                        continue
                    keywords = str(preset.get("keywords", "")).strip()
                    naics = preset.get("naics", [])
                    if isinstance(naics, str):
                        naics = [item.strip() for item in naics.split(",") if item.strip()]
                    if name.strip() and keywords:
                        cleaned[name.strip()] = {
                            "keywords": keywords,
                            "naics": [str(item).strip() for item in naics if str(item).strip()],
                        }
                if cleaned:
                    return cleaned
        except Exception:
            pass
    TRADE_PRESETS_FILE.write_text(
        json.dumps(DEFAULT_TRADE_PRESETS, indent=2),
        encoding="utf-8",
    )
    return dict(DEFAULT_TRADE_PRESETS)


def trade_search_specs(
    selected: list[str],
    presets: dict[str, dict[str, Any]],
    custom_keywords: str = "",
    custom_naics: str = "",
) -> list[dict[str, Any]]:
    """
    Build focused ZoomInfo searches.

    The prior version passed a long comma-separated keyword phrase as both
    industryKeywords and companyDescription, together with all selected NAICS
    codes. ZoomInfo treated that combination too narrowly. Each preset now gets
    its own concise label, its own NAICS codes and optional broader keywords for
    fallback verification.
    """
    specs: list[dict[str, Any]] = []
    for name in selected:
        preset = presets.get(name, {})
        keywords = str(preset.get("keywords", "")).strip()
        codes = preset.get("naics", [])
        if isinstance(codes, str):
            codes = [item.strip() for item in codes.split(",") if item.strip()]
        specs.append({
            "name": name,
            "primary_keyword": name,
            "expanded_keywords": keywords,
            "naics": ",".join(str(code).strip() for code in codes if str(code).strip()),
        })

    if custom_keywords.strip():
        specs.append({
            "name": "Custom",
            "primary_keyword": custom_keywords.split(",")[0].strip(),
            "expanded_keywords": custom_keywords.strip(),
            "naics": custom_naics.strip(),
        })

    return specs


def trade_configuration(
    selected: list[str],
    presets: dict[str, dict[str, Any]],
    custom_keywords: str = "",
    custom_naics: str = "",
) -> tuple[list[str], str]:
    keyword_groups: list[str] = []
    codes: list[str] = []
    for trade in selected:
        preset = presets.get(trade)
        if not preset:
            continue
        keyword_groups.append(str(preset["keywords"]))
        codes.extend(str(code) for code in preset["naics"])
    if custom_keywords.strip():
        keyword_groups.append(custom_keywords.strip())
    if custom_naics.strip():
        codes.extend(
            item.strip() for item in custom_naics.split(",") if item.strip()
        )
    deduped_codes = list(dict.fromkeys(codes))
    return keyword_groups, ",".join(deduped_codes)


def concise_outreach_angle(prospect: "Prospect") -> str:
    pieces = []
    if prospect.growth_signals:
        pieces.append(f"growth activity ({prospect.growth_signals})")
    if prospect.marketing_maturity:
        pieces.append(f"visible marketing investment ({prospect.marketing_maturity})")
    if prospect.service_area_evidence:
        pieces.append(f"defined service territory ({prospect.service_area_evidence})")
    if prospect.technology_signals:
        pieces.append(f"modern operating stack ({prospect.technology_signals})")
    if not pieces:
        return (
            f"{prospect.company_name} appears to be a residential home-service operator "
            "that may benefit from targeted new-mover customer acquisition."
        )
    return (
        f"Lead with {pieces[0]}"
        + (f" and {pieces[1]}" if len(pieces) > 1 else "")
        + "; connect that signal to Darwill's new-mover and measurable customer-acquisition programs."
    )


def normalize_trade_phrase(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (value or "").lower()).strip()


def suggest_trade_mapping(trade_name: str) -> dict[str, Any]:
    normalized = normalize_trade_phrase(trade_name)
    if not normalized:
        return {
            "keywords": "",
            "naics": [],
            "confidence": 0,
            "note": "Enter a trade name.",
        }

    if normalized in TRADE_SUGGESTION_LIBRARY:
        return dict(TRADE_SUGGESTION_LIBRARY[normalized])

    # Token-based fuzzy matching for natural phrases.
    tokens = set(normalized.split())
    best_name = ""
    best_score = 0.0
    for name in TRADE_SUGGESTION_LIBRARY:
        candidate_tokens = set(name.split())
        overlap = len(tokens & candidate_tokens)
        union = len(tokens | candidate_tokens) or 1
        score = overlap / union
        if normalized in name or name in normalized:
            score += 0.5
        if score > best_score:
            best_score = score
            best_name = name

    if best_name and best_score >= 0.45:
        result = dict(TRADE_SUGGESTION_LIBRARY[best_name])
        result["confidence"] = min(int(result.get("confidence", 60)), int(best_score * 100))
        result["note"] = (
            f"Closest known mapping: {best_name}. "
            + str(result.get("note", "Review before saving."))
        )
        return result

    keywords = (
        f"residential {trade_name.strip()}, home {trade_name.strip()}, "
        f"{trade_name.strip()} service, {trade_name.strip()} installation"
    )
    return {
        "keywords": keywords,
        "naics": [],
        "confidence": 35,
        "note": (
            "No strong NAICS match was found automatically. "
            "You may save this as a keyword-only trade or review a classification."
        ),
    }


def load_learning_rules() -> dict[str, Any]:
    default = {
        "rejection_counts": {},
        "blocked_domains": [],
        "blocked_company_terms": [],
        "updated_at": "",
    }
    if not LEARNING_FILE.exists():
        LEARNING_FILE.write_text(json.dumps(default, indent=2), encoding="utf-8")
        return default
    try:
        data = json.loads(LEARNING_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            default.update(data)
    except Exception:
        pass
    return default


def save_learning_rules(rules: dict[str, Any]):
    rules["updated_at"] = now_iso()
    LEARNING_FILE.write_text(json.dumps(rules, indent=2), encoding="utf-8")


MASTER_CSV_COLUMNS = [
    "Prospect Key",
    "ZoomInfo Company ID",
    "Company",
    "Domain",
    "Website",
    "City",
    "State",
    "Trade",
    "Revenue",
    "Employees",
    "Industry",
    "Lifecycle Status",
    "Qualification Decision",
    "Qualification Reason",
    "Fit Score",
    "First Seen",
    "Last Reviewed",
    "Next Review",
    "Approved At",
    "HubSpot Company ID",
    "Sequence Status",
    "Meeting Status",
    "Customer Status",
    "Source Profile",
    "Notes",
]


def master_csv_row_value(
    row: dict[str, Any],
    aliases: list[str],
) -> str:
    normalized = {
        normalize_csv_header(str(key)): value
        for key, value in row.items()
    }
    for alias in aliases:
        value = normalized.get(normalize_csv_header(alias))
        if value not in (None, ""):
            return str(value).strip()
    return ""


def detect_master_csv_candidates() -> list[Path]:
    candidates: list[Path] = []
    search_roots = [
        BASE_DIR.parent,
        Path.home() / "Documents",
        Path.home() / "OneDrive - Darwill",
        Path.home() / "OneDrive",
        Path.home() / "Desktop",
    ]
    names = [
        DEFAULT_MASTER_CSV_NAME,
        DEFAULT_MASTER_CSV_NAME.lower(),
    ]
    for root in search_roots:
        try:
            if not root.exists():
                continue
            for name in names:
                direct = root / name
                if direct.exists():
                    candidates.append(direct)
            # Limit recursive discovery to likely Darwill folders.
            for match in root.glob("**/Darwill_Master_Prospect_Database.csv"):
                candidates.append(match)
        except Exception:
            continue

    unique: list[Path] = []
    seen = set()
    for path in candidates:
        try:
            resolved = str(path.resolve()).lower()
        except Exception:
            resolved = str(path).lower()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(path)
    unique.sort(
        key=lambda path: path.stat().st_mtime if path.exists() else 0,
        reverse=True,
    )
    return unique


def backup_master_csv(path: Path) -> Path | None:
    if not path.exists():
        return None
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup = MASTER_CSV_BACKUP_DIR / (
        f"{path.stem}_{stamp}{path.suffix}"
    )
    shutil.copy2(path, backup)
    return backup


def internal_master_rows(db: "HistoryDB") -> list[tuple]:
    return db.conn.execute(
        """SELECT prospect_key, company_id, company_name, domain, website,
                  city, state, trade, revenue, employees, industry,
                  lifecycle_status, qualification_decision,
                  qualification_reason, fit_score, first_seen_at,
                  last_reviewed_at, next_review_at, approved_at,
                  hubspot_company_id, sequence_status, meeting_status,
                  customer_status, source_profile, notes
           FROM master_prospects
           ORDER BY company_name"""
    ).fetchall()


def sync_internal_master_to_csv(
    db: "HistoryDB",
    path: Path,
    create_backup: bool = True,
) -> dict[str, Any]:
    path.parent.mkdir(parents=True, exist_ok=True)
    backup_path = None
    if create_backup and path.exists():
        backup_path = backup_master_csv(path)

    rows = internal_master_rows(db)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with temp_path.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.writer(handle)
        writer.writerow(MASTER_CSV_COLUMNS)
        writer.writerows(rows)
    temp_path.replace(path)
    return {
        "path": str(path),
        "records": len(rows),
        "backup": str(backup_path) if backup_path else "",
        "updated_at": now_iso(),
    }


def import_master_csv_into_internal_db(
    db: "HistoryDB",
    path: Path,
) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Master CSV not found: {path}")

    rows_read = imported = updated = skipped = 0
    with path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
        errors="replace",
    ) as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            rows_read += 1
            company_name = master_csv_row_value(
                row, ["Company", "Company Name", "company_name"]
            )
            website = master_csv_row_value(
                row, ["Website", "Company Website", "website"]
            )
            domain = master_csv_row_value(
                row, ["Domain", "Company Domain", "domain"]
            )
            company_id = master_csv_row_value(
                row,
                [
                    "ZoomInfo Company ID",
                    "Company ID",
                    "company_id",
                ],
            )
            if not company_name and not domain and not company_id:
                skipped += 1
                continue

            prospect_key = master_csv_row_value(
                row, ["Prospect Key", "prospect_key"]
            )
            if not prospect_key:
                prospect_key = db.prospect_key(
                    company_id,
                    website or domain,
                    company_name,
                )

            existing = db.master_prospect(
                company_id,
                website or domain,
                company_name,
            )
            first_seen = master_csv_row_value(
                row, ["First Seen", "first_seen_at"]
            ) or now_iso()
            updated_at = now_iso()

            values = (
                prospect_key,
                company_id,
                company_name,
                normalize_domain(domain or website),
                website,
                master_csv_row_value(row, ["City", "city"]),
                master_csv_row_value(row, ["State", "state"]),
                master_csv_row_value(row, ["Trade", "trade"]),
                parse_number(
                    master_csv_row_value(row, ["Revenue", "revenue"])
                ),
                parse_number(
                    master_csv_row_value(
                        row, ["Employees", "employees"]
                    )
                ),
                master_csv_row_value(row, ["Industry", "industry"]),
                master_csv_row_value(
                    row,
                    ["Lifecycle Status", "Status", "lifecycle_status"],
                ) or "Qualified",
                master_csv_row_value(
                    row,
                    [
                        "Qualification Decision",
                        "Decision",
                        "qualification_decision",
                    ],
                ),
                master_csv_row_value(
                    row,
                    [
                        "Qualification Reason",
                        "Primary Reason",
                        "qualification_reason",
                    ],
                ),
                float(
                    master_csv_row_value(
                        row, ["Fit Score", "fit_score"]
                    ) or 0
                ),
                first_seen,
                master_csv_row_value(
                    row, ["Last Reviewed", "last_reviewed_at"]
                ),
                master_csv_row_value(
                    row, ["Next Review", "next_review_at"]
                ),
                master_csv_row_value(
                    row, ["Approved At", "approved_at"]
                ),
                master_csv_row_value(
                    row,
                    ["HubSpot Company ID", "hubspot_company_id"],
                ),
                master_csv_row_value(
                    row, ["Sequence Status", "sequence_status"]
                ),
                master_csv_row_value(
                    row, ["Meeting Status", "meeting_status"]
                ),
                master_csv_row_value(
                    row, ["Customer Status", "customer_status"]
                ),
                None,
                master_csv_row_value(
                    row, ["Source Profile", "source_profile"]
                ),
                master_csv_row_value(row, ["Notes", "notes"]),
                updated_at,
            )

            db.conn.execute(
                """INSERT INTO master_prospects(
                       prospect_key, company_id, company_name, domain,
                       website, city, state, trade, revenue, employees,
                       industry, lifecycle_status,
                       qualification_decision, qualification_reason,
                       fit_score, first_seen_at, last_reviewed_at,
                       next_review_at, approved_at, hubspot_company_id,
                       sequence_status, meeting_status, customer_status,
                       last_run_id, source_profile, notes, updated_at
                   ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(prospect_key) DO UPDATE SET
                       company_id=excluded.company_id,
                       company_name=excluded.company_name,
                       domain=excluded.domain,
                       website=excluded.website,
                       city=excluded.city,
                       state=excluded.state,
                       trade=excluded.trade,
                       revenue=excluded.revenue,
                       employees=excluded.employees,
                       industry=excluded.industry,
                       lifecycle_status=CASE
                           WHEN master_prospects.lifecycle_status
                                IN ('Approved','Synced to HubSpot',
                                    'In Sequence','Meeting',
                                    'Customer','Lost')
                           THEN master_prospects.lifecycle_status
                           ELSE excluded.lifecycle_status
                       END,
                       qualification_decision=CASE
                           WHEN excluded.qualification_decision<>''
                           THEN excluded.qualification_decision
                           ELSE master_prospects.qualification_decision
                       END,
                       qualification_reason=CASE
                           WHEN excluded.qualification_reason<>''
                           THEN excluded.qualification_reason
                           ELSE master_prospects.qualification_reason
                       END,
                       fit_score=CASE
                           WHEN excluded.fit_score>0
                           THEN excluded.fit_score
                           ELSE master_prospects.fit_score
                       END,
                       first_seen_at=CASE
                           WHEN master_prospects.first_seen_at<>''
                           THEN master_prospects.first_seen_at
                           ELSE excluded.first_seen_at
                       END,
                       last_reviewed_at=CASE
                           WHEN excluded.last_reviewed_at<>''
                           THEN excluded.last_reviewed_at
                           ELSE master_prospects.last_reviewed_at
                       END,
                       next_review_at=CASE
                           WHEN excluded.next_review_at<>''
                           THEN excluded.next_review_at
                           ELSE master_prospects.next_review_at
                       END,
                       approved_at=CASE
                           WHEN excluded.approved_at<>''
                           THEN excluded.approved_at
                           ELSE master_prospects.approved_at
                       END,
                       hubspot_company_id=CASE
                           WHEN excluded.hubspot_company_id<>''
                           THEN excluded.hubspot_company_id
                           ELSE master_prospects.hubspot_company_id
                       END,
                       sequence_status=CASE
                           WHEN excluded.sequence_status<>''
                           THEN excluded.sequence_status
                           ELSE master_prospects.sequence_status
                       END,
                       meeting_status=CASE
                           WHEN excluded.meeting_status<>''
                           THEN excluded.meeting_status
                           ELSE master_prospects.meeting_status
                       END,
                       customer_status=CASE
                           WHEN excluded.customer_status<>''
                           THEN excluded.customer_status
                           ELSE master_prospects.customer_status
                       END,
                       source_profile=CASE
                           WHEN excluded.source_profile<>''
                           THEN excluded.source_profile
                           ELSE master_prospects.source_profile
                       END,
                       notes=CASE
                           WHEN excluded.notes<>''
                           THEN excluded.notes
                           ELSE master_prospects.notes
                       END,
                       updated_at=excluded.updated_at""",
                values,
            )
            if existing:
                updated += 1
            else:
                imported += 1

    db.conn.commit()
    return {
        "path": str(path),
        "rows_read": rows_read,
        "imported": imported,
        "updated": updated,
        "skipped": skipped,
        "total_internal": db.master_counts().get("Total", 0),
        "imported_at": now_iso(),
    }


def load_outreach_queue() -> list[ReviewQueueItem]:
    if not OUTREACH_QUEUE_FILE.exists():
        return []
    try:
        data = json.loads(OUTREACH_QUEUE_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        valid_fields = set(ReviewQueueItem.__dataclass_fields__)
        items = []
        for row in data:
            if not isinstance(row, dict):
                continue
            clean = {key: value for key, value in row.items() if key in valid_fields}
            defaults = {
                "contact_outreach_order": 0,
                "contact_outreach_order_label": "",
                "contact_source_type": "",
                "contact_source_url": "",
                "contact_email_status": "",
                "contact_email_confidence": 0,
                "contact_phone_status": "",
                "contact_phone_confidence": 0,
                "contact_decision_confidence": 0,
                "contact_data_status": "",
                "contact_recommendation_reason": "",
                "contact_research_summary": "",
                "contact_research_sources": "",
                "contact_public_company_phone": "",
                "contact_predicted_email_pattern": "",
                "contact_acquisition_method": "",
                "contact_acquisition_report": "",
                "contact_email_verification_status": "",
                "contact_email_recovery_method": "",
                "contact_pattern_support_count": 0,
                "contact_domain_mail_status": "",
                "contact_domain_mail_detail": "",
                "contact_deep_recovery_sources": "",
                "contact_zoominfo_email_availability": "Unknown",
                "contact_zoominfo_email_availability_detail": "",
                "contact_zoominfo_enrichment_attempted": False,
                "contact_zoominfo_enrichment_result": "",
                "contact_zoominfo_match_status": "",
                "contact_zoominfo_retry_message": "",
                "contact_zoominfo_enriched_company_id": "",
                "contact_zoominfo_enrichment_warnings": "",
            }
            for key, value in defaults.items():
                clean.setdefault(key, value)
            items.append(ReviewQueueItem(**clean))
        return items
    except Exception:
        return []


def save_outreach_queue(items: list[ReviewQueueItem]):
    OUTREACH_QUEUE_FILE.write_text(
        json.dumps([asdict(item) for item in items], indent=2),
        encoding="utf-8",
    )


def queue_items_from_run(
    prospects: list["Prospect"],
    contacts: list["RankedContact"],
    outreach: list["OutreachDraft"],
) -> list[ReviewQueueItem]:
    prospects_by_id = {item.company_id: item for item in prospects}
    contacts_by_id = {item.contact_id: item for item in contacts}
    existing = {item.queue_id: item for item in load_outreach_queue()}
    output = list(existing.values())

    for draft in outreach:
        prospect = prospects_by_id.get(draft.company_id)
        contact = contacts_by_id.get(draft.contact_id)
        if not prospect or not contact:
            continue
        queue_id = f"{draft.company_id}:{draft.contact_id}"
        if queue_id in existing:
            continue
        item = ReviewQueueItem(
            queue_id=queue_id,
            company_id=draft.company_id,
            company_name=draft.company_name,
            company_website=prospect.website,
            company_state=prospect.state,
            company_city=prospect.city,
            company_revenue=prospect.revenue,
            company_employees=prospect.employees,
            contact_id=draft.contact_id,
            contact_name=draft.contact_name,
            contact_first_name=contact.first_name,
            contact_last_name=contact.last_name,
            contact_title=draft.contact_title,
            contact_email=draft.contact_email,
            contact_phone=contact.direct_phone or contact.mobile_phone,
            contact_rank=draft.contact_rank,
            contact_outreach_order=contact.outreach_order,
            contact_outreach_order_label=contact.outreach_order_label,
            contact_source_type=contact.source_type,
            contact_source_url=contact.source_url,
            contact_email_status=contact.email_status,
            contact_email_confidence=contact.email_confidence,
            contact_phone_status=contact.phone_status,
            contact_phone_confidence=contact.phone_confidence,
            contact_decision_confidence=contact.decision_maker_confidence,
            contact_data_status=contact.contact_data_status,
            contact_recommendation_reason=contact.recommendation_reason,
            contact_research_summary=contact.live_research_summary,
            contact_research_sources=contact.research_sources,
            contact_public_company_phone=contact.public_company_phone,
            contact_predicted_email_pattern=contact.predicted_email_pattern,
            contact_acquisition_method=contact.acquisition_method,
            contact_acquisition_report=contact.acquisition_report,
            contact_email_verification_status=contact.email_verification_status,
            contact_email_recovery_method=contact.email_recovery_method,
            contact_pattern_support_count=contact.pattern_support_count,
            contact_domain_mail_status=contact.domain_mail_status,
            contact_domain_mail_detail=contact.domain_mail_detail,
            contact_deep_recovery_sources=contact.deep_recovery_sources,
            contact_zoominfo_email_availability=(
                contact.zoominfo_email_availability
            ),
            contact_zoominfo_email_availability_detail=(
                contact.zoominfo_email_availability_detail
            ),
            contact_zoominfo_enrichment_attempted=(
                contact.zoominfo_enrichment_attempted
            ),
            contact_zoominfo_enrichment_result=(
                contact.zoominfo_enrichment_result
            ),
            contact_zoominfo_match_status=(
                contact.zoominfo_match_status
            ),
            contact_zoominfo_retry_message=(
                contact.zoominfo_retry_message
            ),
            contact_zoominfo_enriched_company_id=(
                contact.zoominfo_enriched_company_id
            ),
            contact_zoominfo_enrichment_warnings=(
                contact.zoominfo_enrichment_warnings
            ),
            recommended_strategy=draft.recommended_strategy,
            outreach_confidence=draft.outreach_confidence,
            subject_line=draft.subject_line,
            email_body=draft.full_email_draft,
            why_company=prospect.acceptance_reason,
            why_contact=draft.why_this_contact,
            evidence=draft.evidence_used,
            sources=draft.source_urls,
            status="Pending Review",
            created_at=now_iso(),
            updated_at=now_iso(),
        )
        analysis = analyze_deliverability(
            item.subject_line,
            item.email_body,
            item.company_name,
            item.contact_first_name,
            item.recommended_strategy,
            load_deliverability_rules(),
        )
        item.inbox_readiness_score = analysis["inbox_readiness_score"]
        item.subject_quality_score = analysis["subject_quality_score"]
        item.personalization_score = analysis["personalization_score"]
        item.content_risk_score = analysis["content_risk_score"]
        item.human_tone_score = analysis["human_tone_score"]
        item.formatting_score = analysis["formatting_score"]
        item.spam_risk = analysis["spam_risk"]
        item.send_ready = analysis["send_ready"]
        item.deliverability_issues = " | ".join(analysis["issues"])
        item.deliverability_recommendations = " | ".join(
            analysis["recommendations"]
        )
        output.append(item)
    save_outreach_queue(output)
    return output


class HubSpotClient:
    """
    Controlled HubSpot integration.

    The app never enrolls a contact automatically. Record synchronization and
    sequence enrollment are separate explicit user actions.
    """

    BASE_URL = "https://api.hubapi.com"

    def __init__(self, token: str):
        self.token = token.strip()
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def _request(self, method: str, path: str, **kwargs) -> dict[str, Any]:
        response = self.session.request(
            method,
            self.BASE_URL + path,
            timeout=60,
            **kwargs,
        )
        if response.status_code >= 400:
            raise RuntimeError(
                f"HubSpot API error {response.status_code}: {response.text[:1800]}"
            )
        if not response.text.strip():
            return {}
        return response.json()

    def test(self) -> dict[str, Any]:
        return self._request(
            "GET",
            "/crm/v3/objects/contacts",
            params={"limit": 1, "properties": "email"},
        )

    def search_company(self, domain: str, name: str) -> str:
        filters = []
        if domain:
            filters.append({
                "propertyName": "domain",
                "operator": "EQ",
                "value": normalize_domain(domain),
            })
        if not filters and name:
            filters.append({
                "propertyName": "name",
                "operator": "EQ",
                "value": name,
            })
        if not filters:
            return ""
        result = self._request(
            "POST",
            "/crm/v3/objects/companies/search",
            json={
                "filterGroups": [{"filters": filters}],
                "properties": ["name", "domain"],
                "limit": 1,
            },
        )
        rows = result.get("results", [])
        return str(rows[0].get("id", "")) if rows else ""

    def create_company(self, item: ReviewQueueItem) -> str:
        properties = {
            "name": item.company_name,
            "domain": normalize_domain(item.company_website),
            "city": item.company_city,
            "state": item.company_state,
            "description": (
                f"Added by Darwill AI Prospector. "
                f"Strategy: {item.recommended_strategy}. "
                f"Confidence: {item.outreach_confidence}%."
            ),
        }
        properties = {
            key: value for key, value in properties.items()
            if value not in (None, "")
        }
        result = self._request(
            "POST",
            "/crm/v3/objects/companies",
            json={"properties": properties},
        )
        return str(result.get("id", ""))

    def search_contact(self, email: str) -> str:
        if not email:
            return ""
        result = self._request(
            "POST",
            "/crm/v3/objects/contacts/search",
            json={
                "filterGroups": [{
                    "filters": [{
                        "propertyName": "email",
                        "operator": "EQ",
                        "value": email,
                    }]
                }],
                "properties": ["email", "firstname", "lastname"],
                "limit": 1,
            },
        )
        rows = result.get("results", [])
        return str(rows[0].get("id", "")) if rows else ""

    def create_contact(self, item: ReviewQueueItem) -> str:
        properties = {
            "email": item.contact_email,
            "firstname": item.contact_first_name,
            "lastname": item.contact_last_name,
            "jobtitle": item.contact_title,
            "phone": item.contact_phone,
            "company": item.company_name,
        }
        properties = {
            key: value for key, value in properties.items()
            if value not in (None, "")
        }
        result = self._request(
            "POST",
            "/crm/v3/objects/contacts",
            json={"properties": properties},
        )
        return str(result.get("id", ""))

    def associate_contact_company(self, contact_id: str, company_id: str):
        # Standard v3 association endpoint. This does not send any email.
        self._request(
            "PUT",
            f"/crm/v3/objects/contacts/{contact_id}/associations/"
            f"companies/{company_id}/contact_to_company",
        )

    def add_note(self, item: ReviewQueueItem):
        if not item.hubspot_contact_id:
            return
        body = (
            "<b>Darwill AI Prospector Review</b><br>"
            f"<b>Recommended strategy:</b> {item.recommended_strategy}<br>"
            f"<b>Confidence:</b> {item.outreach_confidence}%<br>"
            f"<b>Why company:</b> {item.why_company}<br>"
            f"<b>Why contact:</b> {item.why_contact}<br>"
            f"<b>Approved subject:</b> {item.subject_line}<br>"
            f"<b>Approved email:</b><br>{item.email_body.replace(chr(10), '<br>')}"
        )
        result = self._request(
            "POST",
            "/crm/v3/objects/notes",
            json={
                "properties": {
                    "hs_timestamp": now_iso(),
                    "hs_note_body": body,
                }
            },
        )
        note_id = str(result.get("id", ""))
        if note_id:
            self._request(
                "PUT",
                f"/crm/v3/objects/notes/{note_id}/associations/"
                f"contacts/{item.hubspot_contact_id}/note_to_contact",
            )

    def sync_item(self, item: ReviewQueueItem) -> ReviewQueueItem:
        company_id = self.search_company(
            item.company_website, item.company_name
        )
        if not company_id:
            company_id = self.create_company(item)

        contact_id = self.search_contact(item.contact_email)
        if not contact_id:
            if not item.contact_email:
                raise RuntimeError(
                    f"{item.contact_name} has no email address. "
                    "The contact cannot be created safely for sequence use."
                )
            contact_id = self.create_contact(item)

        self.associate_contact_company(contact_id, company_id)
        item.hubspot_company_id = company_id
        item.hubspot_contact_id = contact_id
        item.hubspot_status = "Synced"
        item.updated_at = now_iso()
        self.add_note(item)
        return item

    def list_sequences(self) -> list[dict[str, Any]]:
        # Try the current date-versioned endpoint, then the v4 endpoint.
        paths = [
            "/automation/sequences/2026-03",
            "/automation/v4/sequences",
        ]
        last_error = None
        for path in paths:
            try:
                result = self._request("GET", path)
                rows = (
                    result.get("results")
                    or result.get("sequences")
                    or result.get("data")
                    or []
                )
                if isinstance(rows, list):
                    return rows
            except Exception as exc:
                last_error = exc
        if last_error:
            raise last_error
        return []

    def enroll_contact(
        self,
        contact_id: str,
        sequence_id: str,
        sender_email: str,
    ) -> dict[str, Any]:
        payload = {
            "contactId": contact_id,
            "sequenceId": sequence_id,
            "senderEmail": sender_email,
        }
        paths = [
            "/automation/sequences/2026-03/enrollments",
            "/automation/v4/sequences/enrollments",
        ]
        last_error = None
        for path in paths:
            try:
                return self._request("POST", path, json=payload)
            except Exception as exc:
                last_error = exc
        if last_error:
            raise last_error
        return {}


def load_darwill_knowledge() -> dict[str, Any]:
    if DARWILL_KNOWLEDGE_FILE.exists():
        try:
            data = json.loads(DARWILL_KNOWLEDGE_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                merged = dict(DEFAULT_DARWILL_KNOWLEDGE)
                merged.update(data)
                merged["offerings"] = {
                    **DEFAULT_DARWILL_KNOWLEDGE["offerings"],
                    **data.get("offerings", {}),
                }
                return merged
        except Exception:
            pass
    DARWILL_KNOWLEDGE_FILE.write_text(
        json.dumps(DEFAULT_DARWILL_KNOWLEDGE, indent=2),
        encoding="utf-8",
    )
    return json.loads(json.dumps(DEFAULT_DARWILL_KNOWLEDGE))


def role_audience(contact: "RankedContact") -> str:
    title = (contact.title or "").lower()
    if any(term in title for term in [
        "marketing", "growth", "brand", "demand generation",
    ]):
        return "marketing"
    if any(term in title for term in [
        "president", "chief executive", "ceo", "owner",
    ]):
        return "executive"
    if any(term in title for term in [
        "chief operating", "coo", "general manager", "operations",
    ]):
        return "operations"
    return "business"


def score_outreach_strategies(
    prospect: "Prospect",
    contact: "RankedContact",
) -> dict[str, int]:
    scores = {name: 35 for name in STRATEGY_NAMES}
    growth = prospect.growth_signals.lower()
    marketing = prospect.marketing_maturity.lower()
    tech = prospect.technology_signals.lower()
    service = prospect.service_area_evidence.lower()
    trade = prospect.trade_signals.lower()
    role = role_audience(contact)
    title = (contact.title or "").lower()

    # New movers
    scores["New Movers"] += 18
    if prospect.residential_signals:
        scores["New Movers"] += 16
    if service:
        scores["New Movers"] += 8
    if any(term in growth for term in ["expansion", "new location"]):
        scores["New Movers"] += 10
    if any(term in trade for term in [
        "hvac", "plumbing", "electrical", "pest", "pool", "garage door",
        "remodel", "roof", "water heater",
    ]):
        scores["New Movers"] += 8

    # Full-service partner
    if marketing:
        scores["Full-Service Marketing Partner"] += 22
    if role == "marketing":
        scores["Full-Service Marketing Partner"] += 14
    if prospect.revenue and prospect.revenue >= 25_000_000:
        scores["Full-Service Marketing Partner"] += 8
    if len([x for x in prospect.trade_signals.split(",") if x.strip()]) >= 3:
        scores["Full-Service Marketing Partner"] += 8

    # Data and attribution
    if any(term in tech for term in [
        "servicetitan", "hubspot", "salesforce", "callrail",
        "google analytics",
    ]):
        scores["Data & Attribution"] += 26
    if role == "marketing":
        scores["Data & Attribution"] += 12
    if any(term in title for term in ["analytics", "performance", "digital"]):
        scores["Data & Attribution"] += 10
    if prospect.revenue and prospect.revenue >= 50_000_000:
        scores["Data & Attribution"] += 8

    # Expansion and awareness
    if any(term in growth for term in [
        "expansion", "new location", "acquisition", "private equity",
    ]):
        scores["Market Expansion & Brand Awareness"] += 30
    if service:
        scores["Market Expansion & Brand Awareness"] += 8
    if role in {"executive", "marketing"}:
        scores["Market Expansion & Brand Awareness"] += 8

    # Retention and reactivation
    if any(term in marketing for term in [
        "membership program", "online booking", "direct mail",
    ]):
        scores["Retention & Reactivation"] += 22
    if any(term in trade for term in [
        "maintenance", "hvac", "pest", "pool", "plumbing",
    ]):
        scores["Retention & Reactivation"] += 10
    if role == "marketing":
        scores["Retention & Reactivation"] += 8

    return {name: min(99, max(1, score)) for name, score in scores.items()}


def strongest_company_signal(prospect: "Prospect") -> str:
    if prospect.growth_signals:
        return f"its current growth activity ({prospect.growth_signals})"
    if prospect.marketing_maturity:
        return f"the visible marketing programs already in place ({prospect.marketing_maturity})"
    if prospect.technology_signals:
        return f"the marketing and operating technology signals found ({prospect.technology_signals})"
    if prospect.service_area_evidence:
        return "the clearly defined service territory and local-market focus"
    if prospect.residential_signals:
        return "the company's strong residential and homeowner focus"
    return "the company's position as a residential home-service provider"


def strategy_reason(
    strategy: str,
    prospect: "Prospect",
    contact: "RankedContact",
) -> str:
    role = contact.title or "business leadership role"
    signal = strongest_company_signal(prospect)
    reasons = {
        "New Movers": (
            f"{prospect.company_name} serves homeowners, and {signal} makes new-household "
            "acquisition a timely and easy-to-understand opening conversation."
        ),
        "Full-Service Marketing Partner": (
            f"{role} is likely to value a broader discussion about how Darwill can support "
            f"multiple acquisition channels, and {signal} suggests the company may benefit "
            "from a partner beyond a single campaign."
        ),
        "Data & Attribution": (
            f"The combination of {role} and {signal} creates a strong reason to lead with "
            "audience intelligence, CRM-informed targeting, and measurable attribution."
        ),
        "Market Expansion & Brand Awareness": (
            f"{signal.capitalize()} indicates a need to build awareness and acquire customers "
            "efficiently across specific markets or service areas."
        ),
        "Retention & Reactivation": (
            f"The company's recurring home-service relationship and {signal} make retention, "
            "reactivation, maintenance-plan growth, and cross-sell relevant."
        ),
    }
    return reasons.get(strategy, signal)


def contact_role_sentence(contact: "RankedContact") -> str:
    first = contact.first_name or "there"
    title = contact.title or "your role"
    role = role_audience(contact)
    if role == "marketing":
        return (
            f"Given your role as {title}, I thought this might be relevant to the way "
            "you approach customer acquisition and campaign performance."
        )
    if role == "executive":
        return (
            f"Given your role as {title}, I thought this might be relevant to the way "
            "you evaluate profitable growth and market share."
        )
    if role == "operations":
        return (
            f"Given your role as {title}, I thought this might be relevant to territory "
            "growth and keeping lead flow aligned with operating capacity."
        )
    return f"Given your role as {title}, I thought this might be worth sharing."


def generate_outreach_draft(
    prospect: "Prospect",
    contact: "RankedContact",
    knowledge: dict[str, Any],
) -> "OutreachDraft":
    scores = score_outreach_strategies(prospect, contact)
    ordered = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    primary_strategy, primary_score = ordered[0]
    secondary_strategy, secondary_score = ordered[1]

    offering = knowledge.get("offerings", {}).get(
        primary_strategy,
        DEFAULT_DARWILL_KNOWLEDGE["offerings"][primary_strategy],
    )
    first_name = contact.first_name or "there"
    company = prospect.company_name
    signal = strongest_company_signal(prospect)
    role_sentence = contact_role_sentence(contact)

    subject_options = {
        "New Movers": [
            f"New homeowner opportunity for {company}",
            f"Territory insight for {company}",
            f"Report I created for {company}",
        ],
        "Full-Service Marketing Partner": [
            f"A broader growth idea for {company}",
            f"Marketing opportunity for {company}",
            f"Thought this might be useful for {company}",
        ],
        "Data & Attribution": [
            f"Audience and attribution idea for {company}",
            f"Measuring customer acquisition at {company}",
            f"Data-driven growth opportunity for {company}",
        ],
        "Market Expansion & Brand Awareness": [
            f"Supporting {company}'s market growth",
            f"Building awareness in {company}'s service areas",
            f"Growth opportunity I noticed for {company}",
        ],
        "Retention & Reactivation": [
            f"Customer reactivation idea for {company}",
            f"Growing repeat business at {company}",
            f"Retention opportunity for {company}",
        ],
    }

    opener = (
        f"Hi {first_name},\n\n"
        f"I was researching {company} and noticed {signal}. {role_sentence}"
    )
    darwill_talking_point = (
        f"{offering.get('value_proposition', '')} "
        f"{offering.get('proof_points', '')}"
    ).strip()
    cta = knowledge.get(
        "meeting_cta",
        DEFAULT_DARWILL_KNOWLEDGE["meeting_cta"],
    )
    full_email = (
        f"{opener}\n\n"
        f"{darwill_talking_point}\n\n"
        f"{cta}\n\n"
        "Thank you,\nJon"
    )

    evidence = " | ".join(
        value for value in [
            prospect.growth_signals,
            prospect.marketing_maturity,
            prospect.technology_signals,
            prospect.service_area_evidence[:300] if prospect.service_area_evidence else "",
            contact.live_research_summary,
        ] if value
    )

    confidence = round(
        min(
            99,
            (primary_score * 0.65)
            + (min(contact.contact_score, 180) / 180 * 25)
            + (10 if prospect.company_sources else 0),
        )
    )

    return OutreachDraft(
        company_id=prospect.company_id,
        company_name=company,
        contact_id=contact.contact_id,
        contact_name=f"{contact.first_name} {contact.last_name}".strip(),
        contact_title=contact.title,
        contact_rank=contact.rank,
        contact_email=contact.email,
        recommended_strategy=primary_strategy,
        strategy_score=primary_score,
        secondary_strategy=secondary_strategy,
        secondary_strategy_score=secondary_score,
        outreach_confidence=confidence,
        subject_line=subject_options[primary_strategy][0],
        alternate_subject_1=subject_options[primary_strategy][1],
        alternate_subject_2=subject_options[primary_strategy][2],
        personalized_opening=opener,
        darwill_talking_point=darwill_talking_point,
        call_to_action=cta,
        full_email_draft=full_email,
        why_this_strategy=strategy_reason(
            primary_strategy, prospect, contact
        ),
        why_this_contact=contact.recommendation_reason,
        evidence_used=evidence,
        source_urls=" | ".join(
            value for value in [
                prospect.company_sources,
                contact.research_sources,
            ] if value
        ),
        new_movers_score=scores["New Movers"],
        full_service_score=scores["Full-Service Marketing Partner"],
        data_attribution_score=scores["Data & Attribution"],
        expansion_awareness_score=scores["Market Expansion & Brand Awareness"],
        retention_reactivation_score=scores["Retention & Reactivation"],
    )


def format_integer_with_commas(value: Any) -> str:
    parsed = parse_number(value)
    return f"{parsed:,}" if parsed is not None else ""


HUBSPOT_COMPANY_HEADER_ALIASES = {
    "record_id": [
        "record id", "company id", "hs object id", "object id",
    ],
    "company_name": [
        "company name", "name", "associated company",
    ],
    "domain": [
        "company domain name", "domain", "website domain",
    ],
    "website": [
        "company website url", "website url", "website",
    ],
    "phone": [
        "phone number", "company phone number", "phone",
    ],
    "city": ["city", "company city"],
    "state": ["state/region", "state", "company state/region"],
    "street": ["street address", "address", "company street address"],
    "postal_code": ["postal code", "zip", "zip code"],
    "owner": ["company owner", "hubspot owner"],
    "last_modified": [
        "last activity date", "last modified date", "updated at",
    ],
}

HUBSPOT_CONTACT_HEADER_ALIASES = {
    "record_id": [
        "record id", "contact id", "hs object id", "object id",
    ],
    "first_name": ["first name", "firstname"],
    "last_name": ["last name", "lastname"],
    "email": ["email", "email address"],
    "phone": ["phone number", "phone"],
    "mobile_phone": ["mobile phone number", "mobile phone"],
    "job_title": ["job title", "jobtitle"],
    "company_name": [
        "associated company", "company name", "associated company name",
    ],
    "company_domain": [
        "associated company domain", "company domain name", "company domain",
    ],
    "linkedin_url": [
        "linkedin url", "linkedin profile", "linkedin",
    ],
    "owner": ["contact owner", "hubspot owner"],
    "last_modified": [
        "last activity date", "last modified date", "updated at",
    ],
}


def normalize_csv_header(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", (value or "").lower())).strip()


def csv_value(
    row: dict[str, Any],
    aliases: list[str],
) -> str:
    normalized = {
        normalize_csv_header(str(key)): value
        for key, value in row.items()
    }
    for alias in aliases:
        value = normalized.get(normalize_csv_header(alias))
        if value not in (None, ""):
            return str(value).strip()
    return ""


def normalized_company_name(value: str) -> str:
    name = normalize_trade_phrase(value)
    suffixes = {
        "llc", "inc", "incorporated", "corp", "corporation",
        "company", "co", "ltd", "limited", "pllc", "lp",
    }
    parts = [part for part in name.split() if part not in suffixes]
    return " ".join(parts)


def normalized_contact_name(first_name: str, last_name: str) -> str:
    return normalize_trade_phrase(f"{first_name} {last_name}")


def normalize_index_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits


def hubspot_csv_kind(fieldnames: list[str]) -> str:
    normalized = {normalize_csv_header(name) for name in fieldnames if name}
    company_signals = {
        "company name", "company domain name", "company website url",
    }
    contact_signals = {
        "first name", "last name", "email",
    }
    company_score = len(normalized & company_signals)
    contact_score = len(normalized & contact_signals)
    if contact_score >= 2:
        return "contacts"
    if company_score >= 1:
        return "companies"
    return "unknown"


def qualification_result(
    prospect: "Prospect", decision: str, reason: str, trade: str
) -> dict[str, Any]:
    return {
        "company_name": prospect.company_name,
        "decision": decision,
        "primary_reason": reason,
        "fit_score": prospect.fit_score,
        "website": prospect.website,
        "city": prospect.city,
        "state": prospect.state,
        "trade": trade,
        "revenue": prospect.revenue,
        "employees": prospect.employees,
        "company_id": prospect.company_id,
        "sources": prospect.company_sources,
        "reviewed_at": now_iso(),
    }


def parse_number(value: Any) -> int | None:
    if value in (None, ""):
        return None
    text = str(value).strip().lower().replace("$", "").replace(",", "")
    multiplier = 1
    if text.endswith("b"):
        multiplier, text = 1_000_000_000, text[:-1]
    elif text.endswith("m"):
        multiplier, text = 1_000_000, text[:-1]
    elif text.endswith("k"):
        multiplier, text = 1_000, text[:-1]
    try:
        return int(float(text) * multiplier)
    except ValueError:
        return None


def crm_linked(record: Any) -> bool:
    blob = text_blob(record).lower()
    positive_patterns = [
        r'"?incrm"?\s*[:=]\s*true',
        r'"?inhubspot"?\s*[:=]\s*true',
        r'"?crm.?match"?\s*[:=]\s*true',
        r'"?exportedtocrm"?\s*[:=]\s*true',
        r'"?crm.?status"?\s*[:=]\s*"?(?:existing|matched|synced|exported)',
        r'hubspot.{0,35}(?:existing|matched|synced|present|exported)',
    ]
    return any(re.search(pattern, blob, re.I) for pattern in positive_patterns)


def result_to_json(result: Any) -> Any:
    if hasattr(result, "model_dump"):
        return result.model_dump()
    if hasattr(result, "dict"):
        return result.dict()
    if isinstance(result, (dict, list, str, int, float, bool)) or result is None:
        return result
    return json.loads(json.dumps(result, default=str))


def decode_json_layers(value: Any, max_depth: int = 5) -> Any:
    current = value
    for _ in range(max_depth):
        if not isinstance(current, str):
            break
        stripped = current.strip()
        if not stripped:
            break
        try:
            current = json.loads(stripped)
        except Exception:
            break
    return current


def extract_tool_payload(result: Any) -> Any:
    data = result_to_json(result)
    if isinstance(data, dict):
        structured = data.get("structuredContent") or data.get("structured_content")
        if structured not in (None, "", [], {}):
            return decode_json_layers(structured)

        content = data.get("content")
        if isinstance(content, list):
            texts = []
            for item in content:
                if isinstance(item, dict):
                    value = item.get("text")
                    if value:
                        texts.append(value)
                elif hasattr(item, "text"):
                    value = getattr(item, "text", "")
                    if value:
                        texts.append(value)
            if texts:
                joined = "\n".join(texts)
                return decode_json_layers(joined)

        if "text" in data and isinstance(data["text"], str):
            return decode_json_layers(data["text"])
    return decode_json_layers(data)


@dataclass
class Prospect:
    company_id: str
    company_name: str
    website: str
    state: str
    city: str
    revenue: int | None
    employees: int | None
    industry: str
    crm_excluded: bool = False
    fit_score: float = 0
    trade_signals: str = ""
    residential_signals: str = ""
    exclusion_signals: str = ""
    company_summary: str = ""
    company_sources: str = ""
    acceptance_reason: str = ""
    growth_signals: str = ""
    marketing_maturity: str = ""
    technology_signals: str = ""
    service_area_evidence: str = ""
    outreach_angle: str = ""


@dataclass
class RankedContact:
    company_id: str
    company_name: str
    contact_id: str
    first_name: str
    last_name: str
    title: str
    email: str
    direct_phone: str
    mobile_phone: str
    linkedin_url: str
    crm_excluded: bool
    rank: str
    contact_score: float
    recommendation_reason: str
    live_research_summary: str
    research_sources: str
    contact_data_status: str = ""
    recommendation_confidence: str = ""
    source_type: str = "ZoomInfo"
    source_url: str = ""
    email_status: str = ""
    phone_status: str = ""
    email_confidence: int = 0
    phone_confidence: int = 0
    decision_maker_confidence: int = 0
    public_company_phone: str = ""
    predicted_email_pattern: str = ""
    outreach_order: int = 0
    outreach_order_label: str = ""
    recommended_cadence: str = ""
    acquisition_method: str = ""
    acquisition_report: str = ""
    email_verification_status: str = ""
    email_recovery_method: str = ""
    pattern_support_count: int = 0
    domain_mail_status: str = ""
    domain_mail_detail: str = ""
    deep_recovery_sources: str = ""
    zoominfo_email_availability: str = "Unknown"
    zoominfo_email_availability_detail: str = ""
    zoominfo_enrichment_attempted: bool = False
    zoominfo_enrichment_result: str = ""
    zoominfo_outer_record_id: str = ""
    zoominfo_person_id_source: str = ""
    zoominfo_match_status: str = ""
    zoominfo_retry_message: str = ""
    zoominfo_enriched_company_id: str = ""
    zoominfo_enrichment_warnings: str = ""


@dataclass
class ReviewQueueItem:
    queue_id: str
    company_id: str
    company_name: str
    company_website: str
    company_state: str
    company_city: str
    company_revenue: int | None
    company_employees: int | None
    contact_id: str
    contact_name: str
    contact_first_name: str
    contact_last_name: str
    contact_title: str
    contact_email: str
    contact_phone: str
    contact_rank: str
    contact_outreach_order: int
    contact_outreach_order_label: str
    contact_source_type: str
    contact_source_url: str
    contact_email_status: str
    contact_email_confidence: int
    contact_phone_status: str
    contact_phone_confidence: int
    contact_decision_confidence: int
    contact_data_status: str
    contact_recommendation_reason: str
    contact_research_summary: str
    contact_research_sources: str
    contact_public_company_phone: str
    contact_predicted_email_pattern: str
    contact_acquisition_method: str
    contact_acquisition_report: str
    contact_email_verification_status: str
    contact_email_recovery_method: str
    contact_pattern_support_count: int
    contact_domain_mail_status: str
    contact_domain_mail_detail: str
    contact_deep_recovery_sources: str
    contact_zoominfo_email_availability: str
    contact_zoominfo_email_availability_detail: str
    contact_zoominfo_enrichment_attempted: bool
    contact_zoominfo_enrichment_result: str
    contact_zoominfo_match_status: str
    contact_zoominfo_retry_message: str
    contact_zoominfo_enriched_company_id: str
    contact_zoominfo_enrichment_warnings: str
    recommended_strategy: str
    outreach_confidence: int
    subject_line: str
    email_body: str
    why_company: str
    why_contact: str
    evidence: str
    sources: str
    status: str = "Pending Review"
    reviewer_notes: str = ""
    hubspot_company_id: str = ""
    hubspot_contact_id: str = ""
    hubspot_status: str = "Not Synced"
    sequence_id: str = ""
    sequence_name: str = ""
    enrollment_status: str = "Not Enrolled"
    inbox_readiness_score: int = 0
    subject_quality_score: int = 0
    personalization_score: int = 0
    content_risk_score: int = 0
    human_tone_score: int = 0
    formatting_score: int = 0
    spam_risk: str = "Not Analyzed"
    send_ready: bool = False
    deliverability_issues: str = ""
    deliverability_recommendations: str = ""
    subject_variant: str = "A"
    outcome: str = "Not Sent"
    created_at: str = ""
    updated_at: str = ""


@dataclass
class OutreachDraft:
    company_id: str
    company_name: str
    contact_id: str
    contact_name: str
    contact_title: str
    contact_rank: str
    contact_email: str
    recommended_strategy: str
    strategy_score: int
    secondary_strategy: str
    secondary_strategy_score: int
    outreach_confidence: int
    subject_line: str
    alternate_subject_1: str
    alternate_subject_2: str
    personalized_opening: str
    darwill_talking_point: str
    call_to_action: str
    full_email_draft: str
    why_this_strategy: str
    why_this_contact: str
    evidence_used: str
    source_urls: str
    new_movers_score: int
    full_service_score: int
    data_attribution_score: int
    expansion_awareness_score: int
    retention_reactivation_score: int


class HistoryDB:
    def __init__(self, path: Path):
        self.conn = sqlite3.connect(path)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS decisions(
                company_id TEXT PRIMARY KEY,
                company_name TEXT,
                decision TEXT,
                reason TEXT,
                updated_at TEXT
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS master_prospects(
                prospect_key TEXT PRIMARY KEY,
                company_id TEXT,
                company_name TEXT,
                domain TEXT,
                website TEXT,
                city TEXT,
                state TEXT,
                trade TEXT,
                revenue INTEGER,
                employees INTEGER,
                industry TEXT,
                lifecycle_status TEXT,
                qualification_decision TEXT,
                qualification_reason TEXT,
                fit_score REAL,
                first_seen_at TEXT,
                last_reviewed_at TEXT,
                next_review_at TEXT,
                approved_at TEXT,
                hubspot_company_id TEXT,
                sequence_status TEXT,
                meeting_status TEXT,
                customer_status TEXT,
                last_run_id INTEGER,
                source_profile TEXT,
                notes TEXT,
                updated_at TEXT
            )
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_master_prospects_company_id
            ON master_prospects(company_id)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_master_prospects_domain
            ON master_prospects(domain)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_master_prospects_status
            ON master_prospects(lifecycle_status)
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS master_contacts(
                contact_key TEXT PRIMARY KEY,
                contact_id TEXT,
                prospect_key TEXT,
                company_id TEXT,
                company_name TEXT,
                first_name TEXT,
                last_name TEXT,
                title TEXT,
                email TEXT,
                phone TEXT,
                linkedin_url TEXT,
                contact_rank TEXT,
                recommendation_score REAL,
                recommendation_reason TEXT,
                crm_status TEXT,
                hubspot_contact_id TEXT,
                sequence_status TEXT,
                outcome TEXT,
                source_type TEXT,
                source_url TEXT,
                email_status TEXT,
                phone_status TEXT,
                email_confidence INTEGER,
                phone_confidence INTEGER,
                decision_maker_confidence INTEGER,
                public_company_phone TEXT,
                predicted_email_pattern TEXT,
                first_seen_at TEXT,
                last_reviewed_at TEXT,
                updated_at TEXT
            )
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_master_contacts_prospect
            ON master_contacts(prospect_key)
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS hubspot_companies_index(
                index_key TEXT PRIMARY KEY,
                hubspot_record_id TEXT,
                company_name TEXT,
                normalized_name TEXT,
                domain TEXT,
                website TEXT,
                phone TEXT,
                normalized_phone TEXT,
                city TEXT,
                state TEXT,
                street TEXT,
                postal_code TEXT,
                owner TEXT,
                source_file TEXT,
                imported_at TEXT,
                last_modified TEXT
            )
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_company_domain
            ON hubspot_companies_index(domain)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_company_name
            ON hubspot_companies_index(normalized_name)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_company_phone
            ON hubspot_companies_index(normalized_phone)
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS hubspot_contacts_index(
                index_key TEXT PRIMARY KEY,
                hubspot_record_id TEXT,
                first_name TEXT,
                last_name TEXT,
                normalized_name TEXT,
                email TEXT,
                phone TEXT,
                normalized_phone TEXT,
                mobile_phone TEXT,
                normalized_mobile_phone TEXT,
                job_title TEXT,
                company_name TEXT,
                normalized_company_name TEXT,
                company_domain TEXT,
                linkedin_url TEXT,
                owner TEXT,
                source_file TEXT,
                imported_at TEXT,
                last_modified TEXT
            )
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_contact_email
            ON hubspot_contacts_index(email)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_contact_name_company
            ON hubspot_contacts_index(normalized_name, normalized_company_name)
        """)
        self.conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_hubspot_contact_phone
            ON hubspot_contacts_index(normalized_phone)
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS hubspot_import_history(
                import_id INTEGER PRIMARY KEY AUTOINCREMENT,
                imported_at TEXT,
                source_file TEXT,
                file_type TEXT,
                import_mode TEXT,
                rows_read INTEGER,
                rows_indexed INTEGER,
                rows_updated INTEGER,
                rows_skipped INTEGER
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS rejection_feedback(
                feedback_id INTEGER PRIMARY KEY AUTOINCREMENT,
                company_id TEXT,
                company_name TEXT,
                website TEXT,
                reason_code TEXT,
                notes TEXT,
                created_at TEXT
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS runs(
                run_id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT,
                completed_at TEXT,
                profile_name TEXT,
                states TEXT,
                trades TEXT,
                qualified_count INTEGER,
                contact_count INTEGER,
                candidate_count INTEGER,
                output_file TEXT,
                status TEXT
            )
        """)
        # Forward-compatible migrations for existing v3.0 databases.
        existing_columns = {
            row[1]
            for row in self.conn.execute(
                "PRAGMA table_info(master_contacts)"
            ).fetchall()
        }
        migrations = {
            "source_type": "TEXT",
            "source_url": "TEXT",
            "email_status": "TEXT",
            "phone_status": "TEXT",
            "email_confidence": "INTEGER",
            "phone_confidence": "INTEGER",
            "decision_maker_confidence": "INTEGER",
            "public_company_phone": "TEXT",
            "predicted_email_pattern": "TEXT",
        }
        for column, column_type in migrations.items():
            if column not in existing_columns:
                self.conn.execute(
                    f"ALTER TABLE master_contacts "
                    f"ADD COLUMN {column} {column_type}"
                )
        self.conn.commit()

    def clear_hubspot_index(self, kind: str):
        if kind in {"companies", "all"}:
            self.conn.execute("DELETE FROM hubspot_companies_index")
        if kind in {"contacts", "all"}:
            self.conn.execute("DELETE FROM hubspot_contacts_index")
        self.conn.commit()

    def import_hubspot_csv(
        self,
        path: Path,
        mode: str = "merge",
    ) -> dict[str, Any]:
        with path.open("r", encoding="utf-8-sig", newline="", errors="replace") as handle:
            reader = csv.DictReader(handle)
            fieldnames = reader.fieldnames or []
            kind = hubspot_csv_kind(fieldnames)
            if kind == "unknown":
                raise RuntimeError(
                    "The CSV could not be identified as a HubSpot company or contact export."
                )
            if mode == "replace":
                self.clear_hubspot_index(kind)

            rows_read = rows_indexed = rows_updated = rows_skipped = 0
            imported_at = now_iso()

            for row in reader:
                rows_read += 1
                if kind == "companies":
                    values = {
                        key: csv_value(row, aliases)
                        for key, aliases in HUBSPOT_COMPANY_HEADER_ALIASES.items()
                    }
                    domain = normalize_domain(values["domain"] or values["website"])
                    normalized_name = normalized_company_name(values["company_name"])
                    normalized_phone = normalize_index_phone(values["phone"])
                    index_key = (
                        f"id:{values['record_id']}"
                        if values["record_id"]
                        else f"domain:{domain}"
                        if domain
                        else f"name:{normalized_name}"
                        if normalized_name
                        else ""
                    )
                    if not index_key:
                        rows_skipped += 1
                        continue
                    exists = self.conn.execute(
                        "SELECT 1 FROM hubspot_companies_index WHERE index_key=?",
                        (index_key,),
                    ).fetchone()
                    self.conn.execute(
                        """INSERT INTO hubspot_companies_index(
                               index_key, hubspot_record_id, company_name,
                               normalized_name, domain, website, phone,
                               normalized_phone, city, state, street,
                               postal_code, owner, source_file, imported_at,
                               last_modified
                           ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                           ON CONFLICT(index_key) DO UPDATE SET
                               hubspot_record_id=excluded.hubspot_record_id,
                               company_name=excluded.company_name,
                               normalized_name=excluded.normalized_name,
                               domain=excluded.domain,
                               website=excluded.website,
                               phone=excluded.phone,
                               normalized_phone=excluded.normalized_phone,
                               city=excluded.city,
                               state=excluded.state,
                               street=excluded.street,
                               postal_code=excluded.postal_code,
                               owner=excluded.owner,
                               source_file=excluded.source_file,
                               imported_at=excluded.imported_at,
                               last_modified=excluded.last_modified""",
                        (
                            index_key, values["record_id"], values["company_name"],
                            normalized_name, domain, values["website"],
                            values["phone"], normalized_phone, values["city"],
                            values["state"], values["street"],
                            values["postal_code"], values["owner"], path.name,
                            imported_at, values["last_modified"],
                        ),
                    )
                else:
                    values = {
                        key: csv_value(row, aliases)
                        for key, aliases in HUBSPOT_CONTACT_HEADER_ALIASES.items()
                    }
                    email = values["email"].strip().lower()
                    normalized_name = normalized_contact_name(
                        values["first_name"], values["last_name"]
                    )
                    company_domain = normalize_domain(values["company_domain"])
                    normalized_company = normalized_company_name(
                        values["company_name"]
                    )
                    normalized_phone = normalize_index_phone(values["phone"])
                    normalized_mobile = normalize_index_phone(values["mobile_phone"])
                    index_key = (
                        f"id:{values['record_id']}"
                        if values["record_id"]
                        else f"email:{email}"
                        if email
                        else (
                            f"name:{normalized_name}|company:"
                            f"{company_domain or normalized_company}"
                        )
                        if normalized_name and (company_domain or normalized_company)
                        else ""
                    )
                    if not index_key:
                        rows_skipped += 1
                        continue
                    exists = self.conn.execute(
                        "SELECT 1 FROM hubspot_contacts_index WHERE index_key=?",
                        (index_key,),
                    ).fetchone()
                    self.conn.execute(
                        """INSERT INTO hubspot_contacts_index(
                               index_key, hubspot_record_id, first_name,
                               last_name, normalized_name, email, phone,
                               normalized_phone, mobile_phone,
                               normalized_mobile_phone, job_title,
                               company_name, normalized_company_name,
                               company_domain, linkedin_url, owner,
                               source_file, imported_at, last_modified
                           ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                           ON CONFLICT(index_key) DO UPDATE SET
                               hubspot_record_id=excluded.hubspot_record_id,
                               first_name=excluded.first_name,
                               last_name=excluded.last_name,
                               normalized_name=excluded.normalized_name,
                               email=excluded.email,
                               phone=excluded.phone,
                               normalized_phone=excluded.normalized_phone,
                               mobile_phone=excluded.mobile_phone,
                               normalized_mobile_phone=excluded.normalized_mobile_phone,
                               job_title=excluded.job_title,
                               company_name=excluded.company_name,
                               normalized_company_name=excluded.normalized_company_name,
                               company_domain=excluded.company_domain,
                               linkedin_url=excluded.linkedin_url,
                               owner=excluded.owner,
                               source_file=excluded.source_file,
                               imported_at=excluded.imported_at,
                               last_modified=excluded.last_modified""",
                        (
                            index_key, values["record_id"], values["first_name"],
                            values["last_name"], normalized_name, email,
                            values["phone"], normalized_phone,
                            values["mobile_phone"], normalized_mobile,
                            values["job_title"], values["company_name"],
                            normalized_company, company_domain,
                            values["linkedin_url"], values["owner"], path.name,
                            imported_at, values["last_modified"],
                        ),
                    )

                if exists:
                    rows_updated += 1
                else:
                    rows_indexed += 1

            self.conn.execute(
                """INSERT INTO hubspot_import_history(
                       imported_at, source_file, file_type, import_mode,
                       rows_read, rows_indexed, rows_updated, rows_skipped
                   ) VALUES(?,?,?,?,?,?,?,?)""",
                (
                    imported_at, path.name, kind, mode, rows_read,
                    rows_indexed, rows_updated, rows_skipped,
                ),
            )
            self.conn.commit()
            return {
                "kind": kind,
                "rows_read": rows_read,
                "rows_indexed": rows_indexed,
                "rows_updated": rows_updated,
                "rows_skipped": rows_skipped,
                "source_file": path.name,
                "imported_at": imported_at,
            }

    def hubspot_company_match(
        self,
        company_name: str,
        website: str,
        phone: str = "",
    ) -> dict[str, Any] | None:
        domain = normalize_domain(website)
        normalized_name = normalized_company_name(company_name)
        normalized_phone = normalize_index_phone(phone)

        if domain:
            row = self.conn.execute(
                """SELECT hubspot_record_id, company_name, domain, phone,
                          city, state, source_file, imported_at
                   FROM hubspot_companies_index
                   WHERE domain=? LIMIT 1""",
                (domain,),
            ).fetchone()
            if row:
                return {
                    "match_type": "Exact domain",
                    "record_id": row[0], "company_name": row[1],
                    "domain": row[2], "phone": row[3], "city": row[4],
                    "state": row[5], "source_file": row[6],
                    "imported_at": row[7],
                }

        if normalized_name:
            rows = self.conn.execute(
                """SELECT hubspot_record_id, company_name, domain, phone,
                          city, state, source_file, imported_at
                   FROM hubspot_companies_index
                   WHERE normalized_name=?""",
                (normalized_name,),
            ).fetchall()
            if len(rows) == 1:
                row = rows[0]
                return {
                    "match_type": "Exact normalized company name",
                    "record_id": row[0], "company_name": row[1],
                    "domain": row[2], "phone": row[3], "city": row[4],
                    "state": row[5], "source_file": row[6],
                    "imported_at": row[7],
                }

        if normalized_phone and len(normalized_phone) >= 10:
            row = self.conn.execute(
                """SELECT hubspot_record_id, company_name, domain, phone,
                          city, state, source_file, imported_at
                   FROM hubspot_companies_index
                   WHERE normalized_phone=? LIMIT 1""",
                (normalized_phone,),
            ).fetchone()
            if row:
                return {
                    "match_type": "Exact company phone",
                    "record_id": row[0], "company_name": row[1],
                    "domain": row[2], "phone": row[3], "city": row[4],
                    "state": row[5], "source_file": row[6],
                    "imported_at": row[7],
                }
        return None

    def hubspot_contact_match(
        self,
        contact: RankedContact,
        prospect: Prospect,
    ) -> dict[str, Any] | None:
        email = (contact.email or "").strip().lower()
        if email and contact.email_status != "Predicted — not verified":
            row = self.conn.execute(
                """SELECT hubspot_record_id, first_name, last_name, email,
                          company_name, company_domain, job_title, source_file,
                          imported_at
                   FROM hubspot_contacts_index
                   WHERE email=? LIMIT 1""",
                (email,),
            ).fetchone()
            if row:
                return {
                    "match_type": "Exact email",
                    "record_id": row[0], "first_name": row[1],
                    "last_name": row[2], "email": row[3],
                    "company_name": row[4], "company_domain": row[5],
                    "job_title": row[6], "source_file": row[7],
                    "imported_at": row[8],
                }

        phone_values = [
            normalize_index_phone(contact.direct_phone),
            normalize_index_phone(contact.mobile_phone),
        ]
        for phone in [value for value in phone_values if len(value) >= 10]:
            row = self.conn.execute(
                """SELECT hubspot_record_id, first_name, last_name, email,
                          company_name, company_domain, job_title, source_file,
                          imported_at
                   FROM hubspot_contacts_index
                   WHERE normalized_phone=? OR normalized_mobile_phone=?
                   LIMIT 1""",
                (phone, phone),
            ).fetchone()
            if row:
                return {
                    "match_type": "Exact phone",
                    "record_id": row[0], "first_name": row[1],
                    "last_name": row[2], "email": row[3],
                    "company_name": row[4], "company_domain": row[5],
                    "job_title": row[6], "source_file": row[7],
                    "imported_at": row[8],
                }

        normalized_name = normalized_contact_name(
            contact.first_name, contact.last_name
        )
        company_domain = normalize_domain(prospect.website)
        company_name = normalized_company_name(prospect.company_name)
        if normalized_name and (company_domain or company_name):
            row = self.conn.execute(
                """SELECT hubspot_record_id, first_name, last_name, email,
                          company_name, company_domain, job_title, source_file,
                          imported_at
                   FROM hubspot_contacts_index
                   WHERE normalized_name=?
                     AND (
                         (?<>'' AND company_domain=?)
                         OR (?<>'' AND normalized_company_name=?)
                     )
                   LIMIT 1""",
                (
                    normalized_name,
                    company_domain, company_domain,
                    company_name, company_name,
                ),
            ).fetchone()
            if row:
                return {
                    "match_type": "Exact name + company",
                    "record_id": row[0], "first_name": row[1],
                    "last_name": row[2], "email": row[3],
                    "company_name": row[4], "company_domain": row[5],
                    "job_title": row[6], "source_file": row[7],
                    "imported_at": row[8],
                }
        return None

    def hubspot_index_stats(self) -> dict[str, Any]:
        company_count = self.conn.execute(
            "SELECT COUNT(*) FROM hubspot_companies_index"
        ).fetchone()[0]
        contact_count = self.conn.execute(
            "SELECT COUNT(*) FROM hubspot_contacts_index"
        ).fetchone()[0]
        last_import = self.conn.execute(
            """SELECT imported_at, source_file, file_type, import_mode,
                      rows_read, rows_indexed, rows_updated, rows_skipped
               FROM hubspot_import_history
               ORDER BY import_id DESC LIMIT 1"""
        ).fetchone()
        return {
            "company_count": int(company_count),
            "contact_count": int(contact_count),
            "last_import": last_import,
        }

    def hubspot_import_history(self, limit: int = 50) -> list[tuple]:
        return self.conn.execute(
            """SELECT imported_at, source_file, file_type, import_mode,
                      rows_read, rows_indexed, rows_updated, rows_skipped
               FROM hubspot_import_history
               ORDER BY import_id DESC LIMIT ?""",
            (limit,),
        ).fetchall()

    def seen(self, company_id: str) -> tuple[str, str] | None:
        row = self.conn.execute(
            "SELECT decision, reason FROM decisions WHERE company_id=?",
            (company_id,),
        ).fetchone()
        return row if row else None

    def save(self, company_id: str, company_name: str, decision: str, reason: str):
        self.conn.execute("""
            INSERT INTO decisions(company_id,company_name,decision,reason,updated_at)
            VALUES(?,?,?,?,?)
            ON CONFLICT(company_id) DO UPDATE SET
              company_name=excluded.company_name,
              decision=excluded.decision,
              reason=excluded.reason,
              updated_at=excluded.updated_at
        """, (company_id, company_name, decision, reason, now_iso()))
        self.conn.commit()

    @staticmethod
    def prospect_key(company_id: str, website: str, company_name: str) -> str:
        if company_id:
            return f"zi:{company_id}"
        domain = normalize_domain(website)
        if domain:
            return f"domain:{domain}"
        normalized_name = normalize_trade_phrase(company_name)
        return f"name:{normalized_name}"

    @staticmethod
    def contact_key(
        contact_id: str,
        email: str,
        first_name: str,
        last_name: str,
        prospect_key: str,
    ) -> str:
        if contact_id:
            return f"zi:{contact_id}"
        if email:
            return f"email:{email.strip().lower()}"
        return (
            f"name:{normalize_trade_phrase(first_name + ' ' + last_name)}"
            f"|{prospect_key}"
        )

    def master_prospect(
        self,
        company_id: str,
        website: str,
        company_name: str,
    ) -> dict[str, Any] | None:
        key = self.prospect_key(company_id, website, company_name)
        row = self.conn.execute(
            """SELECT prospect_key, company_id, company_name, domain, website,
                      city, state, trade, revenue, employees, industry,
                      lifecycle_status, qualification_decision,
                      qualification_reason, fit_score, first_seen_at,
                      last_reviewed_at, next_review_at, approved_at,
                      hubspot_company_id, sequence_status, meeting_status,
                      customer_status, last_run_id, source_profile, notes,
                      updated_at
               FROM master_prospects
               WHERE prospect_key=?""",
            (key,),
        ).fetchone()
        if not row and company_id:
            row = self.conn.execute(
                """SELECT prospect_key, company_id, company_name, domain, website,
                          city, state, trade, revenue, employees, industry,
                          lifecycle_status, qualification_decision,
                          qualification_reason, fit_score, first_seen_at,
                          last_reviewed_at, next_review_at, approved_at,
                          hubspot_company_id, sequence_status, meeting_status,
                          customer_status, last_run_id, source_profile, notes,
                          updated_at
                   FROM master_prospects WHERE company_id=? LIMIT 1""",
                (company_id,),
            ).fetchone()
        if not row:
            domain = normalize_domain(website)
            if domain:
                row = self.conn.execute(
                    """SELECT prospect_key, company_id, company_name, domain, website,
                              city, state, trade, revenue, employees, industry,
                              lifecycle_status, qualification_decision,
                              qualification_reason, fit_score, first_seen_at,
                              last_reviewed_at, next_review_at, approved_at,
                              hubspot_company_id, sequence_status, meeting_status,
                              customer_status, last_run_id, source_profile, notes,
                              updated_at
                       FROM master_prospects WHERE domain=? LIMIT 1""",
                    (domain,),
                ).fetchone()
        if not row:
            return None
        columns = [
            "prospect_key", "company_id", "company_name", "domain", "website",
            "city", "state", "trade", "revenue", "employees", "industry",
            "lifecycle_status", "qualification_decision",
            "qualification_reason", "fit_score", "first_seen_at",
            "last_reviewed_at", "next_review_at", "approved_at",
            "hubspot_company_id", "sequence_status", "meeting_status",
            "customer_status", "last_run_id", "source_profile", "notes",
            "updated_at",
        ]
        return dict(zip(columns, row))

    def should_skip_master(
        self,
        company_id: str,
        website: str,
        company_name: str,
        rejection_review_days: int,
        qualified_review_days: int,
    ) -> tuple[bool, str]:
        record = self.master_prospect(company_id, website, company_name)
        if not record:
            return False, ""

        status = record.get("lifecycle_status") or ""
        reviewed = record.get("last_reviewed_at") or ""
        next_review = record.get("next_review_at") or ""

        if status in {
            "Approved", "Synced to HubSpot", "In Sequence",
            "Meeting", "Customer", "Lost",
        }:
            return True, f"Master database status: {status}"

        if status == "Qualified":
            if qualified_review_days <= 0:
                return True, "Already qualified in master database"
            try:
                due = datetime.fromisoformat(next_review)
                if due.tzinfo is None:
                    due = due.replace(tzinfo=timezone.utc)
                if datetime.now(timezone.utc) < due:
                    return True, (
                        f"Already qualified; re-review eligible {due.date()}"
                    )
            except Exception:
                return True, "Already qualified in master database"

        if status == "Rejected":
            if rejection_review_days <= 0:
                return True, "Previously rejected; re-review disabled"
            try:
                due = datetime.fromisoformat(next_review)
                if due.tzinfo is None:
                    due = due.replace(tzinfo=timezone.utc)
                if datetime.now(timezone.utc) < due:
                    return True, (
                        f"Previously rejected; re-review eligible {due.date()}"
                    )
            except Exception:
                pass

        return False, ""

    def upsert_master_prospect(
        self,
        prospect: Prospect,
        lifecycle_status: str,
        decision: str,
        reason: str,
        trade: str,
        profile_name: str,
        run_id: int | None,
        rereview_days: int = 0,
        notes: str = "",
    ) -> str:
        key = self.prospect_key(
            prospect.company_id, prospect.website, prospect.company_name
        )
        existing = self.master_prospect(
            prospect.company_id, prospect.website, prospect.company_name
        )
        now = now_iso()
        first_seen = (
            existing.get("first_seen_at")
            if existing and existing.get("first_seen_at")
            else now
        )
        next_review = ""
        if rereview_days > 0:
            next_review = (
                datetime.now(timezone.utc) + timedelta(days=rereview_days)
            ).isoformat()
        approved_at = (
            now if lifecycle_status == "Approved"
            else (existing.get("approved_at", "") if existing else "")
        )
        self.conn.execute(
            """INSERT INTO master_prospects(
                   prospect_key, company_id, company_name, domain, website,
                   city, state, trade, revenue, employees, industry,
                   lifecycle_status, qualification_decision,
                   qualification_reason, fit_score, first_seen_at,
                   last_reviewed_at, next_review_at, approved_at,
                   hubspot_company_id, sequence_status, meeting_status,
                   customer_status, last_run_id, source_profile, notes,
                   updated_at
               ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(prospect_key) DO UPDATE SET
                   company_id=excluded.company_id,
                   company_name=excluded.company_name,
                   domain=excluded.domain,
                   website=excluded.website,
                   city=excluded.city,
                   state=excluded.state,
                   trade=excluded.trade,
                   revenue=excluded.revenue,
                   employees=excluded.employees,
                   industry=excluded.industry,
                   lifecycle_status=excluded.lifecycle_status,
                   qualification_decision=excluded.qualification_decision,
                   qualification_reason=excluded.qualification_reason,
                   fit_score=excluded.fit_score,
                   last_reviewed_at=excluded.last_reviewed_at,
                   next_review_at=excluded.next_review_at,
                   approved_at=CASE
                       WHEN excluded.approved_at<>'' THEN excluded.approved_at
                       ELSE master_prospects.approved_at
                   END,
                   last_run_id=excluded.last_run_id,
                   source_profile=excluded.source_profile,
                   notes=CASE
                       WHEN excluded.notes<>'' THEN excluded.notes
                       ELSE master_prospects.notes
                   END,
                   updated_at=excluded.updated_at""",
            (
                key,
                prospect.company_id,
                prospect.company_name,
                normalize_domain(prospect.website),
                prospect.website,
                prospect.city,
                prospect.state,
                trade,
                prospect.revenue,
                prospect.employees,
                prospect.industry,
                lifecycle_status,
                decision,
                reason,
                prospect.fit_score,
                first_seen,
                now,
                next_review,
                approved_at,
                existing.get("hubspot_company_id", "") if existing else "",
                existing.get("sequence_status", "") if existing else "",
                existing.get("meeting_status", "") if existing else "",
                existing.get("customer_status", "") if existing else "",
                run_id,
                profile_name,
                notes,
                now,
            ),
        )
        self.conn.commit()
        return key

    def upsert_master_contact(
        self,
        contact: RankedContact,
        prospect_key: str,
    ) -> str:
        key = self.contact_key(
            contact.contact_id,
            contact.email,
            contact.first_name,
            contact.last_name,
            prospect_key,
        )
        existing = self.conn.execute(
            "SELECT first_seen_at FROM master_contacts WHERE contact_key=?",
            (key,),
        ).fetchone()
        now = now_iso()
        first_seen = existing[0] if existing and existing[0] else now
        self.conn.execute(
            """INSERT INTO master_contacts(
                   contact_key, contact_id, prospect_key, company_id,
                   company_name, first_name, last_name, title, email,
                   phone, linkedin_url, contact_rank,
                   recommendation_score, recommendation_reason,
                   crm_status, hubspot_contact_id, sequence_status,
                   outcome, source_type, source_url, email_status,
                   phone_status, email_confidence, phone_confidence,
                   decision_maker_confidence, public_company_phone,
                   predicted_email_pattern, first_seen_at, last_reviewed_at,
                   updated_at
               ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(contact_key) DO UPDATE SET
                   contact_id=excluded.contact_id,
                   prospect_key=excluded.prospect_key,
                   company_id=excluded.company_id,
                   company_name=excluded.company_name,
                   first_name=excluded.first_name,
                   last_name=excluded.last_name,
                   title=excluded.title,
                   email=excluded.email,
                   phone=excluded.phone,
                   linkedin_url=excluded.linkedin_url,
                   contact_rank=excluded.contact_rank,
                   recommendation_score=excluded.recommendation_score,
                   recommendation_reason=excluded.recommendation_reason,
                   crm_status=excluded.crm_status,
                   source_type=excluded.source_type,
                   source_url=excluded.source_url,
                   email_status=excluded.email_status,
                   phone_status=excluded.phone_status,
                   email_confidence=excluded.email_confidence,
                   phone_confidence=excluded.phone_confidence,
                   decision_maker_confidence=excluded.decision_maker_confidence,
                   public_company_phone=excluded.public_company_phone,
                   predicted_email_pattern=excluded.predicted_email_pattern,
                   last_reviewed_at=excluded.last_reviewed_at,
                   updated_at=excluded.updated_at""",
            (
                key,
                contact.contact_id,
                prospect_key,
                contact.company_id,
                contact.company_name,
                contact.first_name,
                contact.last_name,
                contact.title,
                contact.email,
                contact.direct_phone or contact.mobile_phone,
                contact.linkedin_url,
                contact.rank,
                contact.contact_score,
                contact.recommendation_reason,
                "Excluded" if contact.crm_excluded else "Clear",
                "",
                "",
                "",
                contact.source_type,
                contact.source_url,
                contact.email_status,
                contact.phone_status,
                contact.email_confidence,
                contact.phone_confidence,
                contact.decision_maker_confidence,
                contact.public_company_phone,
                contact.predicted_email_pattern,
                first_seen,
                now,
                now,
            ),
        )
        self.conn.commit()
        return key

    def update_master_lifecycle(
        self,
        company_id: str,
        website: str,
        company_name: str,
        status: str,
        hubspot_company_id: str = "",
        sequence_status: str = "",
        meeting_status: str = "",
        customer_status: str = "",
        notes: str = "",
    ):
        record = self.master_prospect(company_id, website, company_name)
        if not record:
            return
        key = record["prospect_key"]
        self.conn.execute(
            """UPDATE master_prospects SET
                   lifecycle_status=?,
                   hubspot_company_id=CASE WHEN ?<>'' THEN ? ELSE hubspot_company_id END,
                   sequence_status=CASE WHEN ?<>'' THEN ? ELSE sequence_status END,
                   meeting_status=CASE WHEN ?<>'' THEN ? ELSE meeting_status END,
                   customer_status=CASE WHEN ?<>'' THEN ? ELSE customer_status END,
                   notes=CASE WHEN ?<>'' THEN ? ELSE notes END,
                   updated_at=?
               WHERE prospect_key=?""",
            (
                status,
                hubspot_company_id, hubspot_company_id,
                sequence_status, sequence_status,
                meeting_status, meeting_status,
                customer_status, customer_status,
                notes, notes,
                now_iso(),
                key,
            ),
        )
        self.conn.commit()

    def master_prospects(
        self,
        status_filter: str = "All",
        search_text: str = "",
        limit: int = 2000,
    ) -> list[tuple]:
        where = []
        params: list[Any] = []
        if status_filter and status_filter != "All":
            where.append("lifecycle_status=?")
            params.append(status_filter)
        if search_text.strip():
            where.append(
                "(company_name LIKE ? OR domain LIKE ? OR state LIKE ? OR trade LIKE ?)"
            )
            term = f"%{search_text.strip()}%"
            params.extend([term, term, term, term])
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        params.append(limit)
        return self.conn.execute(
            f"""SELECT prospect_key, company_name, domain, city, state, trade,
                       revenue, employees, lifecycle_status,
                       qualification_decision, fit_score, last_reviewed_at,
                       next_review_at, hubspot_company_id, sequence_status,
                       qualification_reason
                FROM master_prospects
                {clause}
                ORDER BY updated_at DESC LIMIT ?""",
            tuple(params),
        ).fetchall()

    def master_counts(self) -> dict[str, int]:
        rows = self.conn.execute(
            """SELECT lifecycle_status, COUNT(*)
               FROM master_prospects GROUP BY lifecycle_status"""
        ).fetchall()
        counts = {status: 0 for status in PROSPECT_STATUSES}
        for status, count in rows:
            counts[status or "New"] = int(count)
        counts["Total"] = sum(counts.values())
        return counts

    def save_rejection_feedback(
        self,
        company_id: str,
        company_name: str,
        website: str,
        reason_code: str,
        notes: str = "",
    ):
        self.conn.execute(
            """INSERT INTO rejection_feedback
               (company_id, company_name, website, reason_code, notes, created_at)
               VALUES(?,?,?,?,?,?)""",
            (
                company_id, company_name, website,
                reason_code, notes, now_iso(),
            ),
        )
        self.conn.commit()

    def rejection_feedback(self, limit: int = 500) -> list[tuple]:
        return self.conn.execute(
            """SELECT company_id, company_name, website, reason_code, notes, created_at
               FROM rejection_feedback
               ORDER BY feedback_id DESC LIMIT ?""",
            (limit,),
        ).fetchall()

    def start_run(self, profile_name: str, states: str, trades: str) -> int:
        cur = self.conn.execute(
            """INSERT INTO runs(started_at, profile_name, states, trades, status)
               VALUES(?,?,?,?,?)""",
            (now_iso(), profile_name, states, trades, "RUNNING"),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def finish_run(
        self,
        run_id: int,
        qualified_count: int,
        contact_count: int,
        candidate_count: int,
        output_file: str,
        status: str,
    ):
        self.conn.execute(
            """UPDATE runs
               SET completed_at=?, qualified_count=?, contact_count=?,
                   candidate_count=?, output_file=?, status=?
               WHERE run_id=?""",
            (
                now_iso(), qualified_count, contact_count, candidate_count,
                output_file, status, run_id,
            ),
        )
        self.conn.commit()

    def recent_runs(self, limit: int = 50) -> list[tuple]:
        return self.conn.execute(
            """SELECT started_at, profile_name, states, trades, qualified_count,
                      contact_count, candidate_count, output_file, status
               FROM runs ORDER BY run_id DESC LIMIT ?""",
            (limit,),
        ).fetchall()


class OAuthManager:
    def __init__(self, client_id: str, client_secret: str, logger):
        self.client_id = client_id.strip()
        self.client_secret = client_secret
        self.logger = logger

    @staticmethod
    def _pkce():
        verifier = base64.urlsafe_b64encode(secrets.token_bytes(64)).rstrip(b"=").decode()
        challenge = base64.urlsafe_b64encode(
            hashlib.sha256(verifier.encode()).digest()
        ).rstrip(b"=").decode()
        return verifier, challenge

    def authorize(self) -> dict[str, Any]:
        verifier, challenge = self._pkce()
        state = secrets.token_urlsafe(32)
        result: dict[str, str] = {}
        event = threading.Event()

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urlparse(self.path)
                query = parse_qs(parsed.query)
                result["code"] = query.get("code", [""])[0]
                result["state"] = query.get("state", [""])[0]
                result["error"] = query.get("error", [""])[0]
                body = (
                    "<html><body style='font-family:Segoe UI;padding:40px'>"
                    "<h2>ZoomInfo authorization received.</h2>"
                    "<p>You can close this browser tab and return to Darwill Prospect Intelligence.</p>"
                    "</body></html>"
                ).encode()
                self.send_response(200)
                self.send_header("Content-Type", "text/html")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                event.set()

            def log_message(self, *_args):
                return

        server = HTTPServer(("localhost", 8080), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()

        params = {
            "client_id": self.client_id,
            "redirect_uri": REDIRECT_URI,
            "response_type": "code",
            "code_challenge": challenge,
            "state": state,
        }
        url = AUTHORIZE_URL + "?" + urlencode(params)
        self.logger("Opening ZoomInfo sign-in and authorization in your browser.")
        webbrowser.open(url)

        if not event.wait(300):
            server.shutdown()
            raise RuntimeError("Authorization timed out after five minutes.")
        server.shutdown()

        if result.get("error"):
            raise RuntimeError(f"ZoomInfo authorization error: {result['error']}")
        if result.get("state") != state:
            raise RuntimeError("OAuth state validation failed.")
        code = result.get("code")
        if not code:
            raise RuntimeError("No authorization code was returned.")

        response = requests.post(
            TOKEN_URL,
            auth=(self.client_id, self.client_secret),
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={
                "grant_type": "authorization_code",
                "code": code,
                "code_verifier": verifier,
                "redirect_uri": REDIRECT_URI,
            },
            timeout=60,
        )
        if response.status_code >= 400:
            raise RuntimeError(
                f"Token exchange failed: HTTP {response.status_code}\n{response.text[:1800]}"
            )
        token = response.json()
        self.store_tokens(token)
        return token

    def refresh(self, refresh_token: str) -> dict[str, Any]:
        response = requests.post(
            TOKEN_URL,
            auth=(self.client_id, self.client_secret),
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
            },
            data={"grant_type": "refresh_token", "refresh_token": refresh_token},
            timeout=60,
        )
        if response.status_code >= 400:
            raise RuntimeError(
                f"Token refresh failed: HTTP {response.status_code}\n{response.text[:1200]}"
            )
        token = response.json()
        if "refresh_token" not in token:
            token["refresh_token"] = refresh_token
        self.store_tokens(token)
        return token

    def _vault_key(self) -> bytes:
        """
        Store only a short Fernet key in Windows Credential Manager.
        OAuth access and refresh tokens can be much larger than Credential
        Manager reliably accepts, so they are encrypted into oauth_tokens.enc.
        """
        existing = keyring.get_password(SERVICE, "oauth_vault_key") or ""
        if existing:
            return existing.encode("ascii")
        key = Fernet.generate_key()
        keyring.set_password(SERVICE, "oauth_vault_key", key.decode("ascii"))
        return key

    def _read_token_vault(self) -> dict[str, Any]:
        if not TOKEN_FILE.exists():
            return {}
        try:
            encrypted = TOKEN_FILE.read_bytes()
            decrypted = Fernet(self._vault_key()).decrypt(encrypted)
            payload = json.loads(decrypted.decode("utf-8"))
            return payload if isinstance(payload, dict) else {}
        except (InvalidToken, ValueError, OSError, json.JSONDecodeError) as exc:
            self.logger(f"Saved OAuth token vault could not be read: {exc}")
            return {}

    def store_tokens(self, token: dict[str, Any]):
        expires_at = int(time.time()) + int(token.get("expires_in", 900)) - 60
        payload = {
            "access_token": token.get("access_token", ""),
            "refresh_token": token.get("refresh_token", ""),
            "expires_at": expires_at,
            "token_type": token.get("token_type", "Bearer"),
            "scope": token.get("scope", ""),
            "saved_at": now_iso(),
        }
        encrypted = Fernet(self._vault_key()).encrypt(
            json.dumps(payload).encode("utf-8")
        )
        TOKEN_FILE.write_bytes(encrypted)

    def clear_token_vault(self):
        try:
            TOKEN_FILE.unlink(missing_ok=True)
        except OSError:
            pass

    def get_access_token(self) -> str:
        saved = self._read_token_vault()
        access = str(saved.get("access_token", "") or "")
        refresh = str(saved.get("refresh_token", "") or "")
        expires_at = int(saved.get("expires_at", 0) or 0)

        if access and time.time() < expires_at:
            return access

        if refresh:
            try:
                return self.refresh(refresh)["access_token"]
            except Exception as exc:
                self.logger(f"Saved authorization could not be refreshed: {exc}")
                self.clear_token_vault()

        return self.authorize()["access_token"]


class ZoomInfoMCP:
    def __init__(self, token: str, logger):
        self.token = token
        self.logger = logger
        self.tools: dict[str, dict[str, Any]] = {}

    def _http_client(self) -> httpx.AsyncClient:
        """
        The MCP Python SDK transport accepts a configured httpx client through
        http_client=. Authentication headers therefore belong on the client,
        not directly on streamable_http_client().
        """
        return httpx.AsyncClient(
            headers={
                "Authorization": f"Bearer {self.token}",
                "Accept": "application/json, text/event-stream",
            },
            follow_redirects=True,
            timeout=httpx.Timeout(120.0, connect=30.0),
        )

    async def discover(self) -> dict[str, dict[str, Any]]:
        async with self._http_client() as custom_client:
            async with streamable_http_client(
                MCP_URL,
                http_client=custom_client,
            ) as streams:
                read_stream, write_stream, _ = streams
                async with ClientSession(read_stream, write_stream) as session:
                    await session.initialize()
                    response = await session.list_tools()
                    tools = response.tools
                    self.tools = {}
                    for tool in tools:
                        dumped = result_to_json(tool)
                        name = dumped.get("name", getattr(tool, "name", ""))
                        self.tools[name] = dumped
                    return self.tools

    async def call(self, tool_name: str, arguments: dict[str, Any]) -> Any:
        async with self._http_client() as custom_client:
            async with streamable_http_client(
                MCP_URL,
                http_client=custom_client,
            ) as streams:
                read_stream, write_stream, _ = streams
                async with ClientSession(read_stream, write_stream) as session:
                    await session.initialize()
                    result = await session.call_tool(tool_name, arguments)
                    return extract_tool_payload(result)

    def run(self, coro):
        return asyncio.run(coro)

    def choose_tool(self, candidates: list[str]) -> str:
        lowered = {name.lower(): name for name in self.tools}
        for candidate in candidates:
            if candidate.lower() in lowered:
                return lowered[candidate.lower()]
        for name in self.tools:
            low = name.lower()
            if any(candidate.lower() in low for candidate in candidates):
                return name
        return ""

    def schema(self, tool_name: str) -> dict[str, Any]:
        tool = self.tools.get(tool_name, {})
        return tool.get("inputSchema") or tool.get("input_schema") or {}

    def build_args(self, tool_name: str, context: dict[str, Any]) -> dict[str, Any]:
        schema = self.schema(tool_name)
        properties = schema.get("properties", {}) if isinstance(schema, dict) else {}
        required = schema.get("required", []) if isinstance(schema, dict) else []
        args: dict[str, Any] = {}

        aliases = {
            "query": context.get("query"),
            "prompt": context.get("query"),
            "criteria": context.get("query"),
            "searchquery": context.get("query"),
            "companyid": context.get("company_id"),
            "companyids": [context.get("company_id")] if context.get("company_id") else None,
            "companyname": context.get("company_name"),
            "companywebsite": context.get("website"),
            "domain": context.get("website"),
            "contactid": context.get("contact_id"),
            "contactids": [context.get("contact_id")] if context.get("contact_id") else None,
            "personid": context.get("contact_id"),
            "personids": [context.get("contact_id")] if context.get("contact_id") else None,
            "salesmotion": "prospecting",
            "motion": "prospecting",
            "limit": context.get("limit", 50),
            "pagesize": context.get("limit", 50),
            "maxresults": context.get("limit", 50),
        }

        for name, spec in properties.items():
            normalized = re.sub(r"[^a-z0-9]", "", name.lower())
            value = aliases.get(normalized)
            if value not in (None, "", []):
                args[name] = value
                continue
            # Common nested filter object.
            if normalized in {"filters", "filter"} and isinstance(spec, dict):
                args[name] = {
                    "query": context.get("query"),
                    "excludeCrm": True,
                    "excludeCRM": True,
                }
            elif normalized in {"excludecrm", "excludeexistingcrm", "excludehubspot"}:
                args[name] = True

        # If schema is open or only requires one unknown string field, use query.
        for required_name in required:
            if required_name not in args:
                spec = properties.get(required_name, {})
                if spec.get("type") == "string":
                    args[required_name] = context.get("query", "")
                elif spec.get("type") == "integer":
                    args[required_name] = context.get("limit", 50)
                elif spec.get("type") == "boolean":
                    args[required_name] = True
                elif spec.get("type") == "array":
                    item_type = (spec.get("items") or {}).get("type")
                    seed = context.get("company_id") or context.get("contact_id")
                    args[required_name] = [seed] if seed else []
                elif spec.get("type") == "object":
                    args[required_name] = {"query": context.get("query", "")}

        if not properties:
            args = {"query": context.get("query", "")}
        return args


def records_from_payload(payload: Any) -> list[dict[str, Any]]:
    payload = decode_json_layers(payload)

    if isinstance(payload, str):
        decoded = decode_json_layers(payload)
        if decoded is not payload:
            return records_from_payload(decoded)
        return []

    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]

    if isinstance(payload, dict):
        # JSON:API data list.
        data = payload.get("data")
        if isinstance(data, list):
            return [item for item in data if isinstance(item, dict)]
        if isinstance(data, dict):
            nested = records_from_payload(data)
            if nested:
                return nested

        for key in [
            "companies", "contacts", "recommendations", "results",
            "records", "items", "matches", "output",
        ]:
            value = payload.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
            if isinstance(value, dict):
                nested = records_from_payload(value)
                if nested:
                    return nested
            if isinstance(value, str):
                nested = records_from_payload(value)
                if nested:
                    return nested

        # Some MCP wrappers put escaped JSON in a text field.
        for key in ["text", "result", "response", "content"]:
            value = payload.get(key)
            if isinstance(value, str):
                nested = records_from_payload(value)
                if nested:
                    return nested

        for value in payload.values():
            if isinstance(value, (dict, list, str)):
                nested = records_from_payload(value)
                if nested:
                    return nested
    return []



ZOOMINFO_EMAIL_AVAILABLE_KEYS = {
    "hasemail", "emailavailable", "isemailavailable",
    "hasbusinessemail", "businessemailavailable",
    "isbusinessemailavailable", "emailisavailable",
}
ZOOMINFO_EMAIL_STATUS_KEYS = {
    "emailstatus", "emailavailability", "emailavailabilitystatus",
    "businessemailstatus", "contactemailstatus",
}


def zoominfo_email_availability(record: dict[str, Any]) -> tuple[str, str]:
    """Inspect non-enriched ZoomInfo data without consuming a credit."""
    if not isinstance(record, dict):
        return "Unknown", "No structured ZoomInfo record was available."

    email = str(first_value(
        record, ["email", "businessEmail", "emailAddress"], ""
    ) or "").strip()
    if email:
        return "Available", "ZoomInfo returned the email in the search result."

    evidence: list[str] = []

    def walk(value: Any, path: str = "") -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                normalized = re.sub(r"[^a-z0-9]", "", str(key).lower())
                child_path = f"{path}.{key}" if path else str(key)
                if (
                    normalized in ZOOMINFO_EMAIL_AVAILABLE_KEYS
                    or normalized in ZOOMINFO_EMAIL_STATUS_KEYS
                ):
                    evidence.append(f"{child_path}={str(child).strip()}")
                walk(child, child_path)
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{path}[{index}]")

    walk(record)
    positive = {"true", "yes", "available", "verified", "1"}
    negative = {"false", "no", "unavailable", "not available", "0"}

    for item in evidence:
        value = item.split("=", 1)[-1].strip().lower()
        if value in positive or "email available" in value:
            return "Available", "ZoomInfo availability evidence: " + item
    for item in evidence:
        value = item.split("=", 1)[-1].strip().lower()
        if value in negative or "email unavailable" in value or "no email" in value:
            return "Unavailable", "ZoomInfo availability evidence: " + item

    blob = text_blob(record).lower()
    if any(p in blob for p in [
        "verified email available", "business email available",
        "email is available",
    ]):
        return "Available", "ZoomInfo metadata indicates an email is available."
    if any(p in blob for p in [
        "no email available", "email unavailable",
        "business email unavailable",
    ]):
        return "Unavailable", "ZoomInfo metadata indicates no email is available."

    return (
        "Unknown",
        "ZoomInfo search results did not expose a definitive email-availability flag.",
    )



def company_from_record(record: dict[str, Any]) -> Prospect:
    attrs = record.get("attributes", record) if isinstance(record, dict) else {}
    company_id = record.get("id", "") if isinstance(record, dict) else ""
    return Prospect(
        company_id=str(
            company_id
            or first_value(attrs, ["companyId", "id", "zoomInfoCompanyId"], "")
        ),
        company_name=str(first_value(attrs, ["companyName", "name"], "")),
        website=str(first_value(attrs, ["website", "domain", "companyWebsite"], "")),
        state=str(first_value(attrs, ["state", "companyState"], "")),
        city=str(first_value(attrs, ["city", "companyCity"], "")),
        revenue=parse_number(first_value(attrs, ["revenue", "annualRevenue"], "")),
        employees=parse_number(first_value(attrs, ["employeeCount", "employees", "headcount"], "")),
        industry=str(first_value(attrs, ["industry", "primaryIndustry", "industries"], "")),
        crm_excluded=crm_linked(record),
    )



EMAIL_FIELD_NAMES = {
    "email", "emailaddress", "businessemail", "workemail",
    "primaryemail", "verifiedemail", "professionalemail",
    "contactemail", "personemail",
}



















def contact_from_record(record: dict[str, Any], company: Prospect) -> RankedContact:
    attrs = record.get("attributes", record) if isinstance(record, dict) else {}
    company_obj = attrs.get("company", {}) if isinstance(attrs, dict) else {}
    contact_id = record.get("id", "") if isinstance(record, dict) else ""

    direct_email = str(
        first_value(
            attrs,
            [
                "email", "businessEmail", "emailAddress", "workEmail",
                "primaryEmail", "verifiedEmail", "professionalEmail",
            ],
            "",
        )
        or ""
    ).strip()
    recursive_email, recursive_path = recursive_email_value(record)
    email = direct_email or recursive_email

    contact = RankedContact(
        company_id=str(
            first_value(attrs, ["companyId", "zoominfoCompanyId"], "")
            or company_obj.get("id", "")
            or company.company_id
        ),
        company_name=str(
            first_value(attrs, ["companyName"], "")
            or company_obj.get("name", "")
            or company.company_name
        ),
        contact_id=str(
            first_value(
                attrs,
                [
                    "personId",
                    "contactId",
                    "zoominfoContactId",
                    "zoomInfoContactId",
                    "id",
                ],
                "",
            )
            or contact_id
        ),
        first_name=str(
            first_value(
                attrs,
                ["firstName", "givenName", "preferredFirstName"],
                "",
            )
        ),
        last_name=str(
            first_value(attrs, ["lastName", "surname", "familyName"], "")
        ),
        title=str(first_value(attrs, ["jobTitle", "title"], "")),
        email=email,
        direct_phone=str(
            first_value(
                attrs,
                ["phone", "directPhone", "directDial", "businessPhone"],
                "",
            )
        ),
        mobile_phone=str(
            first_value(attrs, ["mobilePhone", "mobile", "cellPhone"], "")
        ),
        linkedin_url=str(
            first_value(
                attrs,
                ["linkedInUrl", "linkedin", "linkedinUrl", "externalUrls"],
                "",
            )
        ),
        crm_excluded=crm_linked(record),
        rank="",
        contact_score=0,
        recommendation_reason="",
        live_research_summary="",
        research_sources="",
        contact_data_status="",
        recommendation_confidence="",
        source_type="ZoomInfo",
        source_url="",
        email_status=("ZoomInfo returned" if email else "Missing"),
        phone_status=(
            "ZoomInfo returned"
            if first_value(
                attrs,
                [
                    "phone", "directPhone", "directDial",
                    "mobilePhone", "mobile",
                ],
                "",
            )
            else "Missing"
        ),
        email_confidence=(98 if email else 0),
        phone_confidence=(
            92
            if first_value(
                attrs,
                [
                    "phone", "directPhone", "directDial",
                    "mobilePhone", "mobile",
                ],
                "",
            )
            else 0
        ),
        decision_maker_confidence=0,
        public_company_phone="",
        predicted_email_pattern="",
    )
    contact.zoominfo_outer_record_id = str(contact_id or "")
    actual_person_id = str(
        first_value(
            attrs,
            [
                "personId",
                "contactId",
                "zoominfoContactId",
                "zoomInfoContactId",
                "id",
            ],
            "",
        )
        or ""
    )
    contact.zoominfo_person_id_source = (
        "attributes.personId/contactId"
        if actual_person_id
        else "outer MCP record id fallback"
    )

    if email and recursive_path:
        contact.email_recovery_method = (
            f"ZoomInfo nested field extraction: {recursive_path}"
        )
        contact.email_verification_status = "ZoomInfo Returned"
    return contact


def score_contact(contact: RankedContact) -> tuple[float, list[str]]:
    title = (contact.title or "").lower()

    hard_exclusions = EXCLUDED_TITLES + [
        "preconstruction", "construction manager", "project manager",
        "service manager", "branch manager", "division manager",
        "field operations", "estimating", "engineering",
    ]
    if any(term in title for term in hard_exclusions):
        return -100, ["Non-target or operational function"]

    marketing_weights = [
        ("chief marketing officer", 150), ("cmo", 150),
        ("vice president of marketing", 145), ("vp marketing", 145),
        ("vice president of growth", 142), ("vp growth", 142),
        ("director of marketing", 138), ("marketing director", 138),
        ("director of growth", 136), ("head of marketing", 136),
        ("vice president of sales and marketing", 132),
        ("vp sales and marketing", 132),
        ("marketing manager", 122), ("demand generation", 120),
        ("growth marketing", 120), ("brand director", 116),
    ]
    executive_fallbacks = [
        ("president", 92), ("chief executive officer", 90), ("ceo", 90),
        ("owner", 88), ("chief operating officer", 78), ("coo", 78),
        ("general manager", 60),
    ]

    score = 0.0
    matched = ""
    tier = ""
    for phrase, weight in marketing_weights:
        if phrase in title:
            score, matched, tier = float(weight), phrase, "marketing/growth"
            break
    if score == 0:
        for phrase, weight in executive_fallbacks:
            if phrase in title:
                score, matched, tier = float(weight), phrase, "executive fallback"
                break

    if score == 0:
        return -25, ["Title is not a target marketing, growth, or executive role"]

    reasons = [f"Role match: {matched}", f"Role tier: {tier}"]
    if contact.email:
        score += 10
        reasons.append("email already returned")
    if contact.direct_phone:
        score += 8
        reasons.append("direct phone already returned")
    if contact.mobile_phone:
        score += 5
        reasons.append("mobile already returned")
    if contact.linkedin_url:
        score += 3
        reasons.append("LinkedIn profile returned")
    if contact.crm_excluded:
        score -= 250
        reasons.append("already represented in CRM")
    return score, reasons


def research_company(tavily: TavilyClient, prospect: Prospect) -> Prospect:
    domain = normalize_domain(prospect.website)
    queries = [
        (
            f"site:{domain} residential homeowners services locations service area "
            f"HVAC plumbing electrical pest pool garage door"
            if domain else
            f'"{prospect.company_name}" residential home services service area'
        ),
        f'"{prospect.company_name}" "{prospect.state}" acquisition expansion new location private equity',
        f'"{prospect.company_name}" marketing advertising direct mail google ads facebook ads',
        f'site:{domain} ServiceTitan Housecall Pro HubSpot Salesforce CallRail Podium'
        if domain else f'"{prospect.company_name}" ServiceTitan Housecall Pro HubSpot Salesforce',
    ]
    results = []
    for query in queries:
        try:
            results.extend(
                tavily.search(
                    query=query,
                    search_depth="advanced",
                    max_results=7,
                ).get("results", [])
            )
        except Exception:
            pass

    official = [
        result for result in results
        if domain and normalize_domain(result.get("url", "")) == domain
    ]
    evidence = official or results
    blob = " ".join(
        f"{r.get('title','')} {r.get('content','')} {r.get('raw_content','') or ''}"
        for r in evidence
    ).lower()

    positive = sorted({term for term in POSITIVE_COMPANY_TERMS if term in blob})
    negative = sorted({term for term in NEGATIVE_COMPANY_TERMS if term in blob})
    residential = sorted({
        term for term in ["residential", "homeowner", "homeowners", "your home", "homes"]
        if term in blob
    })

    growth_hits = []
    for label, terms in GROWTH_EVENT_TERMS.items():
        if any(term in blob for term in terms):
            growth_hits.append(label)

    marketing_hits = []
    for label, terms in MARKETING_MATURITY_TERMS.items():
        if any(term in blob for term in terms):
            marketing_hits.append(label)

    tech_hits = []
    for label, terms in TECHNOLOGY_TERMS.items():
        if any(term in blob for term in terms):
            tech_hits.append(label)

    service_area_snippets = []
    for result in evidence:
        content = result.get("content", "")
        lowered = content.lower()
        if any(term in lowered for term in [
            "service area", "areas we serve", "serving", "locations", "counties",
            "communities", "zip code", "cities",
        ]):
            service_area_snippets.append(content[:240])
        if len(service_area_snippets) >= 3:
            break

    score = 0.0
    reasons = []
    if official and positive and residential:
        score += 55
        reasons.append("official-site residential home-service evidence")
    elif positive and residential:
        score += 35
        reasons.append("supporting residential home-service evidence")
    if prospect.revenue and prospect.revenue >= 25_000_000:
        score += 18
        reasons.append("revenue $25M+")
    elif prospect.revenue and prospect.revenue >= 10_000_000:
        score += 12
        reasons.append("revenue $10M+")
    if prospect.employees and 20 <= prospect.employees <= 999:
        score += 10
        reasons.append("target employee scale")
    if growth_hits:
        score += min(10, len(growth_hits) * 3)
        reasons.append("current growth signals")
    if marketing_hits:
        score += min(8, len(marketing_hits) * 2)
        reasons.append("visible marketing maturity")
    if negative:
        score -= 55
        reasons.append("excluded business-model evidence")

    prospect.fit_score = max(score, 0)
    prospect.trade_signals = ", ".join(positive)
    prospect.residential_signals = ", ".join(residential)
    prospect.exclusion_signals = ", ".join(negative)
    prospect.growth_signals = ", ".join(growth_hits)
    prospect.marketing_maturity = ", ".join(marketing_hits)
    prospect.technology_signals = ", ".join(tech_hits)
    prospect.service_area_evidence = " | ".join(service_area_snippets)
    prospect.company_summary = " | ".join(
        r.get("content", "")[:350] for r in evidence[:4]
    )
    prospect.company_sources = " | ".join(
        dict.fromkeys(r.get("url", "") for r in evidence if r.get("url"))
    )
    prospect.acceptance_reason = " | ".join(reasons)
    prospect.outreach_angle = concise_outreach_angle(prospect)
    return prospect



DEEP_CONTACT_PAGE_HINTS = [
    "team", "leadership", "about", "staff", "people", "management",
    "contact", "news", "press", "media", "blog", "company",
]

EMAIL_PATTERN_NAMES = {
    "first.last": "firstname.lastname",
    "firstlast": "firstnamelastname",
    "flast": "first initial + lastname",
    "firstl": "firstname + last initial",
    "last.first": "lastname.firstname",
    "lastf": "lastname + first initial",
    "first": "firstname",
    "last": "lastname",
}


def domain_accepts_mail(domain: str) -> tuple[bool, str]:
    domain = normalize_domain(domain)
    if not domain:
        return False, "No domain"
    try:
        answers = dns.resolver.resolve(domain, "MX", lifetime=5)
        records = sorted(str(answer.exchange).rstrip(".") for answer in answers)
        if records:
            return True, "MX records: " + ", ".join(records[:3])
    except Exception as mx_error:
        try:
            answers = dns.resolver.resolve(domain, "A", lifetime=5)
            if answers:
                return True, (
                    "No MX record returned, but an A record exists; "
                    "mail acceptance is possible but not confirmed."
                )
        except Exception:
            pass
        return False, f"No usable mail DNS record: {mx_error}"
    return False, "No usable mail DNS record"


def fetch_public_page(url: str, timeout: int = 15) -> tuple[str, str]:
    if not url:
        return "", ""
    response = requests.get(
        url,
        timeout=timeout,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/124 Safari/537.36"
            )
        },
    )
    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "").lower()
    if "pdf" in content_type or url.lower().endswith(".pdf"):
        temp_path = BASE_DIR / "_contact_recovery.pdf"
        temp_path.write_bytes(response.content)
        try:
            reader = PdfReader(str(temp_path))
            return "\n".join(page.extract_text() or "" for page in reader.pages[:30]), "PDF"
        finally:
            try:
                temp_path.unlink()
            except Exception:
                pass
    soup = BeautifulSoup(response.text, "html.parser")
    for node in soup(["script", "style", "noscript", "svg"]):
        node.decompose()
    return " ".join(soup.stripped_strings), "HTML"


def discover_candidate_pages(
    tavily: TavilyClient,
    prospect: Prospect,
    contact: RankedContact,
) -> list[str]:
    domain = normalize_domain(prospect.website)
    full_name = f"{contact.first_name} {contact.last_name}".strip()
    queries = [
        f'site:{domain} "{full_name}"' if domain else "",
        f'site:{domain} filetype:pdf "{full_name}"' if domain else "",
        f'"{full_name}" "{prospect.company_name}" email',
        f'"{full_name}" "{prospect.company_name}" contact',
        f'"{full_name}" "{prospect.company_name}" pdf',
        f'"{full_name}" "{prospect.company_name}" press release',
    ]
    urls = []
    for query in queries:
        if not query:
            continue
        try:
            for result in tavily.search(
                query=query, search_depth="advanced", max_results=10
            ).get("results", []):
                if result.get("url"):
                    urls.append(result["url"])
        except Exception:
            pass
    if prospect.website:
        base = prospect.website
        if not re.match(r"https?://", base, re.I):
            base = "https://" + base
        for hint in DEEP_CONTACT_PAGE_HINTS:
            urls.append(urljoin(base.rstrip("/") + "/", hint))
    deduped, seen = [], set()
    for url in urls:
        normalized = url.split("#", 1)[0]
        if normalized and normalized not in seen:
            seen.add(normalized)
            deduped.append(normalized)
    return deduped[:35]


def person_specific_public_email(
    emails: list[str],
    contact: RankedContact,
    company_domain: str,
) -> str:
    first = re.sub(r"[^a-z]", "", contact.first_name.lower())
    last = re.sub(r"[^a-z]", "", contact.last_name.lower())
    candidates = []
    for email in emails:
        email = email.lower().strip()
        if "@" not in email:
            continue
        local, domain = email.split("@", 1)
        if normalize_domain(domain) != company_domain:
            continue
        if local in LOW_VALUE_PUBLIC_EMAIL_PREFIXES:
            continue
        score = (3 if first and first in local else 0) + (4 if last and last in local else 0)
        if score:
            candidates.append((score, email))
    return max(candidates, default=(0, ""))[1]


def collect_domain_email_examples(page_texts: list[str], domain: str) -> list[str]:
    examples = []
    for page_text in page_texts:
        emails, _ = extract_public_contacts(page_text)
        for email in emails:
            if normalize_domain(email.split("@")[-1]) == domain:
                if email_local_part(email) not in LOW_VALUE_PUBLIC_EMAIL_PREFIXES:
                    examples.append(email.lower())
    return sorted(set(examples))


def pattern_from_email_examples(
    emails: list[str],
    known_contacts: list[RankedContact],
    domain: str,
) -> tuple[str, float, int]:
    synthetic = list(known_contacts)
    for known in known_contacts:
        if known.email:
            continue
        first = re.sub(r"[^a-z]", "", known.first_name.lower())
        last = re.sub(r"[^a-z]", "", known.last_name.lower())
        if not first or not last:
            continue
        patterns = {
            "first.last": f"{first}.{last}",
            "firstlast": f"{first}{last}",
            "flast": f"{first[:1]}{last}",
            "firstl": f"{first}{last[:1]}",
            "last.first": f"{last}.{first}",
            "lastf": f"{last}{first[:1]}",
            "first": first,
            "last": last,
        }
        for email in emails:
            if email_local_part(email) in patterns.values():
                clone = RankedContact(**asdict(known))
                clone.email = email
                synthetic.append(clone)
                break
    pattern, confidence = email_pattern_from_known_contacts(synthetic, domain)
    support = 0
    if pattern:
        for known in synthetic:
            if known.email and predict_email(
                known.first_name, known.last_name, domain, pattern
            ) == known.email.lower():
                support += 1
    return pattern, confidence, support


def deep_contact_data_recovery(
    tavily: TavilyClient,
    prospect: Prospect,
    contact: RankedContact,
    known_contacts: list[RankedContact],
    logger,
) -> RankedContact:
    domain = company_domain_from_prospect(prospect)
    if not domain:
        contact.email_verification_status = "Needs Verification"
        contact.domain_mail_status = "No company domain"
        return contact

    page_texts, successful_sources, all_emails, all_phones = [], [], [], []
    for url in discover_candidate_pages(tavily, prospect, contact):
        try:
            page_text, page_kind = fetch_public_page(url)
            if not page_text:
                continue
            lower = page_text.lower()
            full_name = f"{contact.first_name} {contact.last_name}".strip().lower()
            if not (
                full_name in lower
                or prospect.company_name.lower() in lower
                or normalize_domain(url) == domain
            ):
                continue
            page_texts.append(page_text)
            successful_sources.append(f"{page_kind}: {url}")
            emails, phones = extract_public_contacts(page_text)
            all_emails.extend(emails)
            all_phones.extend(phones)
        except Exception:
            continue

    public_email = person_specific_public_email(
        sorted(set(all_emails)), contact, domain
    )
    if public_email:
        contact.email = public_email
        contact.email_status = "Publicly listed"
        contact.email_confidence = 96
        contact.email_verification_status = "Publicly Verified"
        contact.email_recovery_method = "Person-specific email published on a public page"

    if not contact.direct_phone and not contact.mobile_phone and all_phones:
        contact.public_company_phone = normalize_phone(all_phones[0])
        contact.phone_status = "Public company phone"
        contact.phone_confidence = max(contact.phone_confidence, 78)

    if not contact.email:
        examples = collect_domain_email_examples(page_texts, domain)
        pattern, confidence, support = pattern_from_email_examples(
            examples, known_contacts, domain
        )
        if pattern and support >= 2:
            predicted = predict_email(
                contact.first_name, contact.last_name, domain, pattern
            )
            if predicted:
                contact.email = predicted
                contact.email_status = "Predicted — not verified"
                contact.email_confidence = min(89, round(confidence * 100))
                contact.predicted_email_pattern = pattern
                contact.email_verification_status = "Needs Verification"
                contact.email_recovery_method = (
                    f"Predicted from {support} supporting "
                    f"{EMAIL_PATTERN_NAMES.get(pattern, pattern)} examples"
                )
                contact.pattern_support_count = support

    accepts_mail, detail = domain_accepts_mail(domain)
    contact.domain_mail_status = "Mail-capable domain" if accepts_mail else "Mail DNS not confirmed"
    contact.domain_mail_detail = detail

    if contact.email_status == "ZoomInfo returned":
        contact.email_verification_status = "ZoomInfo Returned"
    elif contact.email_status == "Publicly listed":
        contact.email_verification_status = "Publicly Verified"
    elif contact.email_status == "Predicted — not verified":
        contact.email_verification_status = "Needs Verification"
    else:
        contact.email_verification_status = "Missing"

    if successful_sources:
        contact.deep_recovery_sources = " | ".join(successful_sources)
        contact.research_sources = " | ".join(dict.fromkeys(
            value for value in [contact.research_sources, *successful_sources] if value
        ))
    contact.acquisition_report = build_contact_acquisition_report(contact)
    logger(
        f"Deep recovery for {contact.first_name} {contact.last_name} at "
        f"{prospect.company_name}: {contact.email_status or 'no email'}; "
        f"{contact.email_verification_status}; {contact.domain_mail_status}."
    )
    return contact


def build_contact_acquisition_report(contact: RankedContact) -> str:
    evidence = contact.source_url or contact.research_sources or "No source URL recorded"
    email = contact.email or "Not found"
    phone = contact.direct_phone or contact.mobile_phone or contact.public_company_phone or "Not found"
    return (
        f"Contact discovered through: {contact.source_type or 'Unknown'}. "
        f"Role evidence: {contact.recommendation_reason or 'No explanation recorded'}. "
        f"Email: {email} — {contact.email_status or 'Missing'} "
        f"({contact.email_confidence}% confidence). "
        f"Phone: {phone} — {contact.phone_status or 'Missing'} "
        f"({contact.phone_confidence}% confidence). "
        f"Email-pattern inference: {contact.predicted_email_pattern or 'Not used'}. "
        f"Pattern support: {contact.pattern_support_count or 0} example(s). "
        f"Verification status: {contact.email_verification_status or 'Unknown'}. "
        f"Recovery method: {contact.email_recovery_method or 'Not used'}. "
        f"Domain mail status: {contact.domain_mail_status or 'Not checked'} "
        f"({contact.domain_mail_detail or 'No detail'}). "
        f"Supporting source(s): {evidence}. "
        f"Deep recovery source(s): {contact.deep_recovery_sources or 'None'}."
    )


def assign_outreach_order(contacts: list[RankedContact]) -> list[RankedContact]:
    ordered = sorted(
        contacts,
        key=lambda c: (c.decision_maker_confidence, c.contact_score, c.email_confidence),
        reverse=True,
    )
    labels = {1: "Primary — Contact First", 2: "Secondary — Contact Second", 3: "Third — Contact Third"}
    cadence = {
        1: "Email first; call on Day 4–5 if there is no response.",
        2: "Contact after the primary contact has not responded.",
        3: "Use as an executive or alternate escalation contact.",
    }
    for index, contact in enumerate(ordered, 1):
        contact.outreach_order = index
        contact.outreach_order_label = labels.get(index, f"Alternate Contact #{index}")
        contact.recommended_cadence = cadence.get(index, "Hold as an alternate contact.")
        if contact.source_type == "ZoomInfo":
            contact.acquisition_method = "Returned by ZoomInfo and evaluated with live public research"
        elif contact.email_status == "Predicted — not verified":
            contact.acquisition_method = (
                f"Found through {contact.source_type}; email predicted from the "
                f"supported {contact.predicted_email_pattern} company pattern"
            )
        else:
            contact.acquisition_method = f"Discovered through {contact.source_type or 'public research'}"
        contact.acquisition_report = build_contact_acquisition_report(contact)
    return ordered


def discover_public_decision_makers(
    tavily: TavilyClient,
    prospect: Prospect,
) -> tuple[list[RankedContact], str]:
    domain = normalize_domain(prospect.website)
    queries = [
        f'site:{domain} leadership marketing growth president owner team'
        if domain else
        f'"{prospect.company_name}" leadership marketing growth president owner',
        f'"{prospect.company_name}" "vice president of marketing"',
        f'"{prospect.company_name}" "director of marketing"',
        f'"{prospect.company_name}" president owner leadership',
        f'"{prospect.company_name}" appointed marketing OR growth OR president',
    ]

    results = []
    for query in queries:
        try:
            results.extend(
                tavily.search(
                    query=query,
                    search_depth="advanced",
                    max_results=8,
                ).get("results", [])
            )
        except Exception:
            pass

    company_blob = " ".join(
        f"{result.get('title','')} {result.get('content','')} "
        f"{result.get('raw_content','') or ''}"
        for result in results
    )
    public_emails, public_phones = extract_public_contacts(company_blob)

    company_phone = ""
    if public_phones:
        company_phone = public_phones[0]

    contacts: dict[str, RankedContact] = {}
    name_title_patterns = [
        re.compile(
            r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\s*[,|—–-]\s*"
            r"((?:Chief|Vice President|VP|Director|Head|President|Owner|CEO|COO|"
            r"General Manager|Marketing Manager|Growth Manager|Chief Revenue Officer)"
            r"[^.;|\n]{0,80})"
        ),
        re.compile(
            r"\b((?:Chief|Vice President|VP|Director|Head|President|Owner|CEO|COO|"
            r"General Manager|Marketing Manager|Growth Manager|Chief Revenue Officer)"
            r"[^.;|\n]{0,70})\s*[:|—–-]\s*"
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})"
        ),
    ]

    for result in results:
        title_text = result.get("title", "")
        content = result.get("content", "")
        raw = result.get("raw_content", "") or ""
        url = result.get("url", "")
        blob = f"{title_text}\n{content}\n{raw}"
        result_domain = normalize_domain(url)
        source_type = (
            "Company Website"
            if domain and result_domain == domain
            else "Press Release"
            if any(term in url.lower() for term in ["news", "press", "prnewswire", "businesswire"])
            else "Public Search"
        )

        found_pairs = []
        for pattern_index, pattern in enumerate(name_title_patterns):
            for match in pattern.finditer(blob):
                if pattern_index == 0:
                    full_name, job_title = match.group(1), match.group(2)
                else:
                    job_title, full_name = match.group(1), match.group(2)
                found_pairs.append((full_name.strip(), job_title.strip()))

        for full_name, job_title in found_pairs:
            if title_relevance_score(job_title) <= 0:
                continue
            parts = full_name.split()
            if len(parts) < 2:
                continue
            first_name = parts[0]
            last_name = parts[-1]
            key = f"{full_name.lower()}|{job_title.lower()}"
            if key in contacts:
                continue

            page_emails, page_phones = extract_public_contacts(blob)
            personal_email = ""
            for email in page_emails:
                local = email_local_part(email)
                if normalize_domain(email.split("@")[-1]) != domain:
                    continue
                if local in LOW_VALUE_PUBLIC_EMAIL_PREFIXES:
                    continue
                first = re.sub(r"[^a-z]", "", first_name.lower())
                last = re.sub(r"[^a-z]", "", last_name.lower())
                if first in local or last in local:
                    personal_email = email
                    break

            public_phone = page_phones[0] if page_phones else company_phone
            base_score = title_relevance_score(job_title)
            source_score = CONTACT_SOURCE_PRIORITY.get(source_type, 60)
            decision_confidence = min(
                98,
                round(base_score * 0.55 + source_score * 0.45),
            )

            contact = RankedContact(
                company_id=prospect.company_id,
                company_name=prospect.company_name,
                contact_id=unique_public_contact_id(
                    prospect.company_id, full_name, job_title
                ),
                first_name=first_name,
                last_name=last_name,
                title=job_title,
                email=personal_email,
                direct_phone="",
                mobile_phone="",
                linkedin_url="",
                crm_excluded=False,
                rank="",
                contact_score=float(base_score),
                recommendation_reason=(
                    f"Public source identifies {full_name} as {job_title}. "
                    f"Source: {source_type}."
                ),
                live_research_summary=(
                    f"Publicly identified decision-maker tied to "
                    f"{prospect.company_name}."
                ),
                research_sources=url,
                contact_data_status=(
                    "Public email found"
                    if personal_email
                    else "No public direct email found"
                ),
                recommendation_confidence=(
                    "High" if decision_confidence >= 85
                    else "Medium" if decision_confidence >= 70
                    else "Review"
                ),
                source_type=source_type,
                source_url=url,
                email_status=(
                    "Publicly listed" if personal_email else "Missing"
                ),
                phone_status=(
                    "Public company phone" if public_phone else "Missing"
                ),
                email_confidence=(90 if personal_email else 0),
                phone_confidence=(75 if public_phone else 0),
                decision_maker_confidence=decision_confidence,
                public_company_phone=public_phone,
                predicted_email_pattern="",
            )
            contacts[key] = contact

    return list(contacts.values()), company_phone


def merge_decision_maker_candidates(
    zoominfo_contacts: list[RankedContact],
    public_contacts: list[RankedContact],
    prospect: Prospect,
) -> list[RankedContact]:
    domain = company_domain_from_prospect(prospect)
    pattern, pattern_confidence = email_pattern_from_known_contacts(
        zoominfo_contacts, domain
    )

    merged: dict[str, RankedContact] = {}
    for contact in zoominfo_contacts + public_contacts:
        name_key = normalize_trade_phrase(
            f"{contact.first_name} {contact.last_name}"
        )
        key = name_key or contact.contact_id
        existing = merged.get(key)
        if existing is None:
            merged[key] = contact
            continue

        # Prefer ZoomInfo identity/contact data while preserving public evidence.
        preferred = existing
        secondary = contact
        if CONTACT_SOURCE_PRIORITY.get(contact.source_type, 0) > CONTACT_SOURCE_PRIORITY.get(existing.source_type, 0):
            preferred, secondary = contact, existing
        if existing.source_type == "ZoomInfo":
            preferred, secondary = existing, contact
        preferred.email = preferred.email or secondary.email
        preferred.direct_phone = preferred.direct_phone or secondary.direct_phone
        preferred.mobile_phone = preferred.mobile_phone or secondary.mobile_phone
        preferred.public_company_phone = (
            preferred.public_company_phone
            or secondary.public_company_phone
        )
        preferred.linkedin_url = (
            preferred.linkedin_url or secondary.linkedin_url
        )
        preferred.research_sources = " | ".join(
            dict.fromkeys(
                value for value in [
                    preferred.research_sources,
                    secondary.research_sources,
                    preferred.source_url,
                    secondary.source_url,
                ] if value
            )
        )
        preferred.recommendation_reason = " ".join(
            value for value in [
                preferred.recommendation_reason,
                secondary.recommendation_reason,
            ] if value
        )
        preferred.decision_maker_confidence = max(
            preferred.decision_maker_confidence,
            secondary.decision_maker_confidence,
        )
        if preferred.email:
            if preferred.source_type == "ZoomInfo":
                preferred.email_status = "ZoomInfo returned"
                preferred.email_confidence = max(
                    preferred.email_confidence, 95
                )
            elif not preferred.email_status:
                preferred.email_status = "Publicly listed"
                preferred.email_confidence = max(
                    preferred.email_confidence, 88
                )
        merged[key] = preferred

    output = []
    for contact in merged.values():
        if not contact.email and pattern and domain:
            predicted = predict_email(
                contact.first_name,
                contact.last_name,
                domain,
                pattern,
            )
            if predicted:
                contact.email = predicted
                contact.email_status = "Predicted — not verified"
                contact.email_confidence = round(
                    pattern_confidence * 100
                )
                contact.predicted_email_pattern = pattern
                contact.contact_data_status = (
                    f"Predicted email using {pattern} pattern; not verified"
                )

        role_score = title_relevance_score(contact.title)
        source_score = CONTACT_SOURCE_PRIORITY.get(
            contact.source_type, 50
        )
        evidence_bonus = 10 if contact.research_sources else 0
        contactability_bonus = (
            8 if contact.email_status in {"ZoomInfo returned", "Publicly listed"}
            else 2 if contact.email_status == "Predicted — not verified"
            else 0
        )
        contact.contact_score = max(
            contact.contact_score,
            role_score + source_score * 0.15
            + evidence_bonus + contactability_bonus,
        )
        if not contact.decision_maker_confidence:
            contact.decision_maker_confidence = min(
                99,
                round(role_score * 0.65 + source_score * 0.35),
            )
        output.append(contact)

    output.sort(
        key=lambda contact: (
            contact.decision_maker_confidence,
            contact.contact_score,
            contact.email_confidence,
        ),
        reverse=True,
    )
    return output


def research_contact(tavily: TavilyClient, contact: RankedContact, company_website: str = "") -> RankedContact:
    full_name = f"{contact.first_name} {contact.last_name}".strip()
    if not full_name:
        contact.live_research_summary = "No contact name available for live research."
        return contact

    domain = normalize_domain(company_website)
    queries = [
        f'"{full_name}" "{contact.company_name}" "{contact.title}"',
        f'site:{domain} "{full_name}"' if domain else "",
    ]
    results = []
    for query in queries:
        if not query:
            continue
        try:
            results.extend(
                tavily.search(
                    query=query,
                    search_depth="advanced",
                    max_results=6,
                ).get("results", [])
            )
        except Exception:
            pass

    tied_results = []
    full_lower = full_name.lower()
    company_lower = contact.company_name.lower()
    for result in results:
        title = result.get("title", "")
        content = result.get("content", "")
        url = result.get("url", "")
        blob = f"{title} {content}".lower()
        result_domain = normalize_domain(url)

        # Count only evidence explicitly tying the named person to the company,
        # or a page on the company's own domain containing the person's name.
        explicitly_tied = full_lower in blob and company_lower in blob
        official_tied = bool(domain and result_domain == domain and full_lower in blob)
        if explicitly_tied or official_tied:
            tied_results.append(result)

    tied_blob = " ".join(
        f"{r.get('title','')} {r.get('content','')}" for r in tied_results
    ).lower()
    responsibility_terms = [
        term for term in [
            "marketing", "growth", "customer acquisition", "brand",
            "demand generation", "advertising", "business development",
            "revenue growth", "marketing strategy",
        ] if term in tied_blob
    ]

    employment_confirmed = bool(tied_results)
    if employment_confirmed:
        contact.contact_score += 15
    if responsibility_terms:
        contact.contact_score += min(18, len(responsibility_terms) * 3)

    contact.live_research_summary = (
        f"Current company connection: {'confirmed' if employment_confirmed else 'not independently confirmed'}. "
        f"Relevant responsibility signals tied to this person/company: "
        f"{', '.join(responsibility_terms) or 'none found'}."
    )
    contact.research_sources = " | ".join(
        dict.fromkeys(r.get("url", "") for r in tied_results if r.get("url"))
    )
    return contact


def export_workbook(path: Path, prospects: list[Prospect], contacts: list[RankedContact], outreach: list[OutreachDraft], decisions: list[dict[str, Any]], summary: dict):
    workbook = xlsxwriter.Workbook(path)
    header = workbook.add_format({
        "bold": True, "font_color": WHITE, "bg_color": BLUE_DARK, "border": 1,
    })
    section = workbook.add_format({
        "bold": True, "font_color": WHITE, "bg_color": BLUE, "border": 1,
    })
    cell = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
    money = workbook.add_format({"num_format": "$#,##0", "border": 1})
    integer = workbook.add_format({"num_format": "#,##0", "border": 1})
    score_fmt = workbook.add_format({"num_format": "0.0", "border": 1})
    link = workbook.add_format({"font_color": BLUE, "underline": 1, "border": 1})

    def write_sheet(name, rows, columns):
        sheet = workbook.add_worksheet(name)
        sheet.freeze_panes(1, 0)
        sheet.set_tab_color(BLUE)
        for j, (label, _, _) in enumerate(columns):
            sheet.write(0, j, label, header)
        for i, row in enumerate(rows, 1):
            for j, (_, key, kind) in enumerate(columns):
                value = row.get(key, "")
                fmt = cell
                if kind == "money":
                    fmt = money
                elif kind == "int":
                    fmt = integer
                elif kind == "score":
                    fmt = score_fmt
                elif kind == "url" and value:
                    url = value if str(value).startswith("http") else "https://" + str(value)
                    sheet.write_url(i, j, url, link, string=str(value))
                    continue
                sheet.write(i, j, value, fmt)
        sheet.autofilter(0, 0, max(len(rows), 1), len(columns) - 1)
        for j, (label, _, _) in enumerate(columns):
            sheet.set_column(j, j, min(max(len(label) + 4, 14), 44))

    prospect_cols = [
        ("Fit Score", "fit_score", "score"), ("Company", "company_name", "text"),
        ("Website", "website", "url"), ("State", "state", "text"),
        ("City", "city", "text"), ("Revenue", "revenue", "money"),
        ("Employees", "employees", "int"), ("Industry", "industry", "text"),
        ("Trade Signals", "trade_signals", "text"),
        ("Residential Signals", "residential_signals", "text"),
        ("Acceptance Reason", "acceptance_reason", "text"),
        ("Growth Signals", "growth_signals", "text"),
        ("Marketing Maturity", "marketing_maturity", "text"),
        ("Technology Signals", "technology_signals", "text"),
        ("Service Area Evidence", "service_area_evidence", "text"),
        ("Personalized Outreach Angle", "outreach_angle", "text"),
        ("Company Summary", "company_summary", "text"),
        ("Sources", "company_sources", "text"), ("ZoomInfo ID", "company_id", "text"),
    ]
    contact_cols = [
        ("Company", "company_name", "text"), ("Rank", "rank", "text"),
        ("Contact Score", "contact_score", "score"), ("First Name", "first_name", "text"),
        ("Last Name", "last_name", "text"), ("Title", "title", "text"),
        ("Email", "email", "text"), ("Direct Phone", "direct_phone", "text"),
        ("Mobile Phone", "mobile_phone", "text"), ("LinkedIn", "linkedin_url", "url"),
        ("Why Recommended", "recommendation_reason", "text"),
        ("Live Research", "live_research_summary", "text"),
        ("Research Sources", "research_sources", "text"),
        ("Contact Data Status", "contact_data_status", "text"),
        ("Recommendation Confidence", "recommendation_confidence", "text"),
        ("Decision-Maker Confidence", "decision_maker_confidence", "int"),
        ("Source Type", "source_type", "text"),
        ("Source URL", "source_url", "text"),
        ("Email Status", "email_status", "text"),
        ("Email Confidence", "email_confidence", "int"),
        ("Phone Status", "phone_status", "text"),
        ("Phone Confidence", "phone_confidence", "int"),
        ("Public Company Phone", "public_company_phone", "text"),
        ("Predicted Email Pattern", "predicted_email_pattern", "text"),
        ("Outreach Order", "outreach_order", "int"),
        ("Outreach Order Label", "outreach_order_label", "text"),
        ("Recommended Cadence", "recommended_cadence", "text"),
        ("Acquisition Method", "acquisition_method", "text"),
        ("Contact Acquisition Report", "acquisition_report", "text"),
        ("Email Verification Status", "email_verification_status", "text"),
        ("Email Recovery Method", "email_recovery_method", "text"),
        ("Pattern Support Count", "pattern_support_count", "int"),
        ("Domain Mail Status", "domain_mail_status", "text"),
        ("Domain Mail Detail", "domain_mail_detail", "text"),
        ("Deep Recovery Sources", "deep_recovery_sources", "text"),
        ("Contact ID", "contact_id", "text"),
    ]
    outreach_cols = [
        ("Company", "company_name", "text"),
        ("Contact", "contact_name", "text"),
        ("Title", "contact_title", "text"),
        ("Contact Rank", "contact_rank", "text"),
        ("Email", "contact_email", "text"),
        ("Recommended Strategy", "recommended_strategy", "text"),
        ("Strategy Score", "strategy_score", "int"),
        ("Secondary Strategy", "secondary_strategy", "text"),
        ("Secondary Score", "secondary_strategy_score", "int"),
        ("Outreach Confidence", "outreach_confidence", "int"),
        ("Subject Line", "subject_line", "text"),
        ("Alternate Subject 1", "alternate_subject_1", "text"),
        ("Alternate Subject 2", "alternate_subject_2", "text"),
        ("Personalized Opening", "personalized_opening", "text"),
        ("Darwill Talking Point", "darwill_talking_point", "text"),
        ("Call to Action", "call_to_action", "text"),
        ("Full Email Draft", "full_email_draft", "text"),
        ("Why This Strategy", "why_this_strategy", "text"),
        ("Why This Contact", "why_this_contact", "text"),
        ("Evidence Used", "evidence_used", "text"),
        ("Sources", "source_urls", "text"),
        ("New Movers Score", "new_movers_score", "int"),
        ("Full-Service Score", "full_service_score", "int"),
        ("Data & Attribution Score", "data_attribution_score", "int"),
        ("Expansion & Awareness Score", "expansion_awareness_score", "int"),
        ("Retention & Reactivation Score", "retention_reactivation_score", "int"),
    ]
    decision_cols = [
        ("Company", "company_name", "text"), ("Decision", "decision", "text"),
        ("Primary Reason", "primary_reason", "text"), ("Fit Score", "fit_score", "score"),
        ("Website", "website", "url"), ("City", "city", "text"),
        ("State", "state", "text"), ("Trade", "trade", "text"),
        ("Revenue", "revenue", "money"), ("Employees", "employees", "int"),
        ("ZoomInfo ID", "company_id", "text"), ("Sources", "sources", "text"),
        ("Reviewed At", "reviewed_at", "text"),
    ]
    acquisition_cols = [
        ("Company", "company_name", "text"), ("Outreach Order", "outreach_order", "int"),
        ("Outreach Order Label", "outreach_order_label", "text"),
        ("First Name", "first_name", "text"), ("Last Name", "last_name", "text"),
        ("Title", "title", "text"), ("Decision-Maker Confidence", "decision_maker_confidence", "int"),
        ("Found Through", "source_type", "text"), ("Source URL", "source_url", "url"),
        ("Acquisition Method", "acquisition_method", "text"), ("Email", "email", "text"),
        ("Email Status", "email_status", "text"), ("Email Confidence", "email_confidence", "int"),
        ("Direct Phone", "direct_phone", "text"), ("Mobile Phone", "mobile_phone", "text"),
        ("Public Company Phone", "public_company_phone", "text"),
        ("Phone Status", "phone_status", "text"), ("Phone Confidence", "phone_confidence", "int"),
        ("Predicted Email Pattern", "predicted_email_pattern", "text"),
        ("Why Recommended", "recommendation_reason", "text"),
        ("Research Summary", "live_research_summary", "text"),
        ("Research Sources", "research_sources", "text"),
        ("Recommended Cadence", "recommended_cadence", "text"),
        ("Full Acquisition Report", "acquisition_report", "text"),
        ("Email Verification Status", "email_verification_status", "text"),
        ("Email Recovery Method", "email_recovery_method", "text"),
        ("Pattern Support Count", "pattern_support_count", "int"),
        ("Domain Mail Status", "domain_mail_status", "text"),
        ("Domain Mail Detail", "domain_mail_detail", "text"),
        ("Deep Recovery Sources", "deep_recovery_sources", "text"),
    ]
    write_sheet("Qualified Companies", [asdict(x) for x in prospects], prospect_cols)
    write_sheet("Qualification Decisions", decisions, decision_cols)
    write_sheet("Best Contacts", [asdict(x) for x in contacts], contact_cols)
    write_sheet("Contact Acquisition", [asdict(x) for x in contacts], acquisition_cols)
    write_sheet("AI Outreach", [asdict(x) for x in outreach], outreach_cols)

    sheet = workbook.add_worksheet("Run Summary")
    sheet.set_tab_color(BLUE_DARK)
    sheet.write("A1", "Darwill AI Prospector — Run Summary", section)
    sheet.set_column("A:A", 38)
    sheet.set_column("B:B", 84)
    row = 2
    for key, value in summary.items():
        sheet.write(row, 0, key.replace("_", " ").title(), header)
        sheet.write(row, 1, str(value), cell)
        row += 1
    workbook.close()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x920")
        self.minsize(1120, 760)
        self.configure(bg=SURFACE_ALT)
        try:
            self.state("zoomed")
        except Exception:
            pass
        self.events = queue.Queue()
        self.stop_requested = False
        self.mcp_tools: dict[str, dict[str, Any]] = {}
        self.run_started_at = 0.0
        self.current_run_id: int | None = None
        self.profile_data: dict[str, dict[str, Any]] = {}
        self.trade_presets = load_trade_presets()
        self.learning_rules = load_learning_rules()
        self.darwill_knowledge = load_darwill_knowledge()
        self.review_queue: list[ReviewQueueItem] = load_outreach_queue()
        self.deliverability_rules = load_deliverability_rules()
        self.intelligence_store = IntelligenceStore(DB_PATH)
        self.hubspot_sequences: list[dict[str, Any]] = []
        self._configure_style()
        self.sidebar_collapsed = False
        self._sidebar_buttons = []
        self._sidebar_section_labels = []
        self.global_search_var = tk.StringVar()
        self.workspace_status_var = tk.StringVar(
            value="Ready · Proven engine preserved"
        )
        self.company_intelligence_vars = {
            "fit": tk.StringVar(value="Not evaluated"),
            "confidence": tk.StringVar(value="—"),
            "email_status": tk.StringVar(value="Not evaluated"),
            "next_action": tk.StringVar(value="Review company"),
            "executive_summary": tk.StringVar(
                value="Select a company to generate an executive summary."
            ),
        }
        self._build()
        self._load_settings()
        self._load_profiles()
        self._render_trade_checkboxes()
        self._refresh_trade_preview()
        self._sync_state_chips()
        self._refresh_connection_status()
        self._update_search_mode()
        self._update_advanced_visibility()
        self._refresh_history()
        self._refresh_learning()
        self._refresh_deal_desk()
        self._refresh_master_database()
        self._refresh_hubspot_workspace()
        self._refresh_hubspot_index()
        self._initialize_permanent_master_csv()
        self.after(100, self._poll)
        self.after(1000, self._tick_elapsed)
        self.after(120, self._show_splash)

    def _configure_style(self):
        style = ttk.Style(self)
        apply_compass_theme(style)
        style.theme_use("clam")

        style.configure(
            ".",
            font=("Segoe UI", 10),
            foreground=TEXT,
            background=SURFACE_ALT,
        )
        style.configure("TFrame", background=SURFACE_ALT)
        style.configure("Card.TFrame", background=SURFACE)
        style.configure("TLabel", background=SURFACE_ALT, foreground=TEXT)
        style.configure("Card.TLabel", background=SURFACE, foreground=TEXT)

        style.configure(
            "ProductTitle.TLabel",
            background=NAVY,
            foreground=WHITE,
            font=("Segoe UI Semibold", 22),
        )
        style.configure(
            "ProductSubtitle.TLabel",
            background=NAVY,
            foreground="#CFE3F6",
            font=("Segoe UI", 10),
        )
        style.configure(
            "SectionTitle.TLabel",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 12),
        )
        style.configure(
            "Muted.TLabel",
            background=SURFACE,
            foreground=MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "PageEyebrow.TLabel",
            background=SURFACE,
            foreground=ACCENT,
            font=("Segoe UI Semibold", 8),
        )
        style.configure(
            "PageTitle.TLabel",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "PageSubtitle.TLabel",
            background=SURFACE,
            foreground=MUTED,
            font=("Segoe UI", 10),
        )
        style.configure(
            "CardHeader.TLabel",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 11),
        )

        style.configure(
            "Primary.TButton",
            background=ACCENT,
            foreground=WHITE,
            font=("Segoe UI Semibold", 10),
            padding=(16, 10),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", ACCENT_HOVER),
                ("pressed", ACCENT_HOVER),
                ("disabled", "#9EB9D2"),
            ],
            foreground=[("disabled", "#EDF4FA")],
        )

        style.configure(
            "Secondary.TButton",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 10),
            padding=(14, 9),
            bordercolor=BORDER,
            borderwidth=1,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", BLUE_LIGHT)],
            bordercolor=[("active", ACCENT)],
        )

        # Preserve legacy style name for existing calls.
        style.configure(
            "Blue.TButton",
            background=ACCENT,
            foreground=WHITE,
            font=("Segoe UI Semibold", 10),
            padding=(14, 9),
            borderwidth=0,
        )
        style.map(
            "Blue.TButton",
            background=[("active", ACCENT_HOVER), ("disabled", "#9EB9D2")],
        )

        style.configure(
            "TLabelframe",
            background=SURFACE,
            bordercolor="#C9D9E7",
            relief="solid",
            borderwidth=1,
            padding=4,
        )
        style.configure(
            "TLabelframe.Label",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 11),
        )

        style.configure(
            "TNotebook",
            background=SURFACE_ALT,
            borderwidth=0,
            tabmargins=(0, 8, 0, 0),
        )
        style.configure(
            "Main.TNotebook",
            background=CONTENT_BG,
            borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        style.layout("Main.TNotebook.Tab", [])
        style.configure(
            "TNotebook.Tab",
            background="#E8F0F7",
            foreground=NAVY,
            padding=(18, 10),
            font=("Segoe UI Semibold", 10),
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", SURFACE), ("active", BLUE_LIGHT)],
            foreground=[("selected", ACCENT)],
        )

        style.configure(
            "Horizontal.TProgressbar",
            background=ACCENT,
            troughcolor="#DCE8F2",
            borderwidth=0,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
        )

        style.configure(
            "Treeview",
            background=SURFACE,
            fieldbackground=SURFACE,
            foreground=TEXT,
            rowheight=30,
            bordercolor=BORDER,
            borderwidth=1,
        )
        style.configure(
            "Treeview.Heading",
            background=NAVY_2,
            foreground=WHITE,
            font=("Segoe UI Semibold", 9),
            padding=(8, 8),
        )
        style.map(
            "Treeview",
            background=[("selected", "#DCEEFF")],
            foreground=[("selected", NAVY)],
        )

        style.configure(
            "TEntry",
            padding=10,
            fieldbackground="#FBFDFF",
            foreground=TEXT,
            bordercolor="#B9CAD9",
            lightcolor="#B9CAD9",
            darkcolor="#B9CAD9",
        )
        style.configure(
            "TCombobox",
            padding=9,
            fieldbackground="#FBFDFF",
            foreground=TEXT,
            bordercolor="#B9CAD9",
            lightcolor="#B9CAD9",
            darkcolor="#B9CAD9",
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#B8C7D5",
            troughcolor="#E7EEF5",
            bordercolor="#E7EEF5",
            arrowcolor=NAVY,
        )

        style.configure(
            "ProfileCard.TLabelframe",
            background=SURFACE,
            bordercolor="#D4E1EC",
            relief="solid",
            borderwidth=1,
            padding=8,
        )
        style.configure(
            "ProfileCard.TLabelframe.Label",
            background=SURFACE,
            foreground=NAVY,
            font=("Segoe UI Semibold", 12),
        )
        style.configure(
            "MetricEntry.TEntry",
            padding=12,
            fieldbackground="#F8FBFE",
            foreground=NAVY,
            font=("Segoe UI Semibold", 12),
            bordercolor="#B9CAD9",
            lightcolor="#B9CAD9",
            darkcolor="#B9CAD9",
        )
        style.configure(
            "Hero.TButton",
            background=ACCENT,
            foreground=WHITE,
            font=("Segoe UI Semibold", 13),
            padding=(24, 14),
            borderwidth=0,
        )
        style.map(
            "Hero.TButton",
            background=[
                ("active", ACCENT_HOVER),
                ("pressed", ACCENT_HOVER),
                ("disabled", "#9EB9D2"),
            ],
        )

    def _build(self):
        # Product header — modern Compass shell, proven 3.6 engine.
        header = tk.Frame(self, bg=SIDEBAR, height=96)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand_wrap = tk.Frame(header, bg=SIDEBAR)
        brand_wrap.pack(side="left", fill="y", padx=(24, 0))

        logo = tk.Canvas(
            brand_wrap,
            width=54,
            height=54,
            bg=SIDEBAR,
            highlightthickness=0,
        )
        logo.pack(side="left", pady=20)
        logo.create_oval(
            3, 3, 51, 51,
            fill=ACCENT,
            outline="#75B6F5",
            width=2,
        )
        logo.create_text(
            27, 28,
            text="D",
            fill=WHITE,
            font=("Segoe UI Semibold", 24),
        )

        title_wrap = tk.Frame(brand_wrap, bg=SIDEBAR)
        title_wrap.pack(side="left", padx=(16, 0), pady=13)
        tk.Label(
            title_wrap,
            text="DARWILL",
            bg=SIDEBAR,
            fg=WHITE,
            font=("Segoe UI Black", 19),
        ).pack(anchor="w")
        tk.Label(
            title_wrap,
            text="Compass",
            bg=SIDEBAR,
            fg="#78B7F4",
            font=("Segoe UI Semibold", 15),
        ).pack(anchor="w", pady=(0, 0))
        tk.Label(
            title_wrap,
            text="Proven 3.6 intelligence engine",
            bg=SIDEBAR,
            fg="#9AB8D2",
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(2, 0))

        right_header = tk.Frame(header, bg=SIDEBAR)
        right_header.pack(side="right", fill="y", padx=(0, 24))
        tk.Label(
            right_header,
            text="VERSION 8.6",
            bg=SIDEBAR,
            fg="#79A9D1",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="e", pady=(20, 5))
        tk.Button(
            right_header,
            text="About",
            command=self._show_about,
            bg=SIDEBAR_SECTION,
            fg=WHITE,
            activebackground=SIDEBAR_HOVER,
            activeforeground=WHITE,
            relief="flat",
            bd=0,
            padx=18,
            pady=7,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="e")

        command_bar = tk.Frame(
            self,
            bg=workspace_palette()["command_bg"],
            padx=14,
            pady=7,
            highlightthickness=1,
            highlightbackground=workspace_palette()["border"],
        )
        command_bar.pack(fill="x")

        tk.Button(
            command_bar,
            text="☰",
            command=self._toggle_sidebar,
            bg=workspace_palette()["command_bg"],
            fg=TEXT,
            activebackground=BLUE_LIGHT,
            activeforeground=TEXT,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2",
            font=("Segoe UI Semibold", 11),
        ).pack(side="left")

        tk.Label(
            command_bar,
            text="WORKSPACE SEARCH",
            bg=workspace_palette()["command_bg"],
            fg=MUTED,
            font=("Segoe UI Semibold", 7),
        ).pack(side="left", padx=(10, 8))

        search_shell = tk.Frame(
            command_bar,
            bg=WHITE,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        search_shell.pack(side="left", fill="x", expand=True)

        search_entry = tk.Entry(
            search_shell,
            textvariable=self.global_search_var,
            relief="flat",
            bd=0,
            bg=WHITE,
            fg=TEXT,
            insertbackground=TEXT,
            font=("Segoe UI", 10),
        )
        search_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=10,
            pady=6,
        )
        search_entry.bind(
            "<Return>",
            lambda _event: self._run_global_workspace_search(),
        )

        tk.Button(
            search_shell,
            text="Search",
            command=self._run_global_workspace_search,
            bg=ACCENT,
            fg=WHITE,
            activebackground=ACCENT_HOVER,
            activeforeground=WHITE,
            relief="flat",
            bd=0,
            padx=14,
            pady=6,
            cursor="hand2",
            font=("Segoe UI Semibold", 8),
        ).pack(side="right")

        tk.Label(
            command_bar,
            textvariable=self.workspace_status_var,
            bg=workspace_palette()["command_bg"],
            fg=SUCCESS,
            font=("Segoe UI Semibold", 8),
        ).pack(side="right", padx=(12, 0))

        shell = tk.Frame(self, bg=CONTENT_BG)
        shell.pack(fill="both", expand=True)

        sidebar = tk.Frame(shell, bg=SIDEBAR, width=230)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        self.sidebar_frame = sidebar
        self.sidebar_expanded_width = 230
        self.sidebar_collapsed_width = 62

        content = ttk.Frame(
            shell,
            style="TFrame",
            padding=(14, 12, 14, 14),
        )
        content.pack(side="right", fill="both", expand=True)
        self.workspace_content = content

        notebook = ttk.Notebook(content, style="Main.TNotebook")
        notebook.pack(fill="both", expand=True)
        self.main_notebook = notebook

        settings_tab = ttk.Frame(notebook, style="Card.TFrame")
        run = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        deal_desk_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        deliverability_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        master_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        hubspot_workspace_tab = ttk.Frame(
            notebook, padding=14, style="Card.TFrame"
        )
        hubspot_index_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        history_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        learning_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")
        knowledge_tab = ttk.Frame(notebook, padding=14, style="Card.TFrame")

        pages = [
            (settings_tab, "Search Profiles", "DISCOVERY"),
            (run, "Run Dashboard", "DISCOVERY"),
            (deal_desk_tab, "Deal Desk", "DISCOVERY"),
            (master_tab, "Master Database", "DATA"),
            (hubspot_workspace_tab, "HubSpot Workspace", "CRM"),
            (hubspot_index_tab, "HubSpot CSV Index", "CRM"),
            (history_tab, "Run History", "DATA"),
            (deliverability_tab, "Deliverability", "INTELLIGENCE"),
            (learning_tab, "Qualification Learning", "INTELLIGENCE"),
            (knowledge_tab, "Darwill Knowledge", "INTELLIGENCE"),
        ]
        for frame, label, _section in pages:
            notebook.add(frame, text=label)

        nav_buttons: dict[int, tk.Button] = {}
        active_index = tk.IntVar(value=0)

        def select_page(index: int):
            notebook.select(index)
            active_index.set(index)
            for button_index, button in nav_buttons.items():
                if button_index == index:
                    button.configure(
                        bg=SIDEBAR_ACTIVE,
                        fg=WHITE,
                        activebackground=SIDEBAR_ACTIVE,
                    )
                else:
                    button.configure(
                        bg=SIDEBAR,
                        fg="#D7E5F2",
                        activebackground=SIDEBAR_HOVER,
                    )
            self.workspace_status_var.set(
                f"{pages[index][1]} · Ready"
            )

        self._select_workspace_page = select_page
        self._workspace_pages = pages

        current_section = None
        for index, (_frame, label, section) in enumerate(pages):
            if section != current_section:
                current_section = section
                section_label = tk.Label(
                    sidebar,
                    text=section,
                    bg=SIDEBAR_SECTION,
                    fg="#7EA3C3",
                    anchor="w",
                    padx=18,
                    pady=7,
                    font=("Segoe UI Semibold", 8),
                )
                section_label.pack(
                    fill="x",
                    pady=(10 if index else 14, 4),
                )
                self._sidebar_section_labels.append(section_label)
            button = tk.Button(
                sidebar,
                text=label,
                command=lambda i=index: select_page(i),
                bg=SIDEBAR,
                fg="#D7E5F2",
                activebackground=SIDEBAR_HOVER,
                activeforeground=WHITE,
                relief="flat",
                bd=0,
                anchor="w",
                padx=22,
                pady=10,
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
            )
            button.pack(fill="x", padx=10, pady=2)
            button._compass_full_text = label
            button._compass_short_text = label[:1].upper()
            self._sidebar_buttons.append(button)
            nav_buttons[index] = button

        footer = tk.Frame(sidebar, bg=SIDEBAR_SECTION)
        footer.pack(side="bottom", fill="x", padx=12, pady=14)
        tk.Label(
            footer,
            text="Darwill Compass 8.6",
            bg=SIDEBAR_SECTION,
            fg=WHITE,
            anchor="w",
            font=("Segoe UI Semibold", 9),
        ).pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(
            footer,
            text="●  Proven engine preserved",
            bg=SIDEBAR_SECTION,
            fg="#6FE0A8",
            anchor="w",
            font=("Segoe UI", 8),
        ).pack(fill="x", padx=12, pady=(0, 10))

        select_page(0)

        settings_canvas = tk.Canvas(
            settings_tab,
            bg=CONTENT_BG,
            highlightthickness=0,
            borderwidth=0,
        )
        settings_scrollbar = ttk.Scrollbar(
            settings_tab,
            orient="vertical",
            command=settings_canvas.yview,
        )
        settings = ttk.Frame(
            settings_canvas,
            padding=14,
            style="Card.TFrame",
        )
        settings_window = settings_canvas.create_window(
            (0, 0),
            window=settings,
            anchor="nw",
        )

        def _resize_settings_content(event):
            settings_canvas.itemconfigure(
                settings_window,
                width=event.width,
            )

        def _update_settings_scrollregion(_event=None):
            settings_canvas.configure(
                scrollregion=settings_canvas.bbox("all")
            )

        settings_canvas.bind("<Configure>", _resize_settings_content)
        settings.bind("<Configure>", _update_settings_scrollregion)
        settings_canvas.configure(yscrollcommand=settings_scrollbar.set)
        settings_canvas.pack(side="left", fill="both", expand=True)
        settings_scrollbar.pack(side="right", fill="y")

        def _settings_mousewheel(event):
            delta = -1 if event.delta > 0 else 1
            settings_canvas.yview_scroll(delta * 3, "units")

        def _bind_settings_wheel(_event):
            settings_canvas.bind_all("<MouseWheel>", _settings_mousewheel)

        def _unbind_settings_wheel(_event):
            settings_canvas.unbind_all("<MouseWheel>")

        settings_canvas.bind("<Enter>", _bind_settings_wheel)
        settings_canvas.bind("<Leave>", _unbind_settings_wheel)

        profile_banner = tk.Frame(
            settings,
            bg=NAVY,
            highlightthickness=0,
        )
        profile_banner.pack(fill="x", pady=(0, 12))
        tk.Label(
            profile_banner,
            text="SEARCH PROFILES",
            bg=NAVY,
            fg="#78B7F4",
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(
            profile_banner,
            text="Configure discovery with the proven 3.6 engine.",
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 15),
        ).pack(anchor="w", padx=18)
        tk.Label(
            profile_banner,
            text="Target companies → verify fit → recover contacts → review and approve",
            bg=NAVY,
            fg="#BCD3E7",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=18, pady=(3, 14))

        self.client_id = tk.StringVar()
        self.client_secret = tk.StringVar()
        self.tavily_key = tk.StringVar()
        self.states = tk.StringVar(value=DEFAULT_STATES)
        self.trades = tk.StringVar(value=DEFAULT_TRADES)
        self.custom_trade_keywords = tk.StringVar(value="")
        self.search_mode = tk.StringVar(value="states")
        self.minimum_revenue = tk.StringVar(value="10,000,000")
        self.maximum_revenue = tk.StringVar(value="250,000,000")
        self.target_count = tk.IntVar(value=20)
        self.candidate_limit = tk.IntVar(value=100)
        self.contacts_per_company = tk.IntVar(value=3)
        self.minimum_fit_score = tk.IntVar(value=50)
        self.naics_codes = tk.StringVar(value="")
        default_selected = {
            "HVAC", "Plumbing", "Electrical", "Pest Control",
            "Pool Service", "Garage Door"
        }
        self.trade_vars = {
            name: tk.BooleanVar(value=name in default_selected)
            for name in self.trade_presets
        }
        self.employee_min = tk.IntVar(value=20)
        self.employee_max = tk.IntVar(value=1000)
        self.exclude_crm = tk.BooleanVar(value=True)
        self.enrich_final_contacts = tk.BooleanVar(value=False)
        self.smart_email_enrichment = tk.BooleanVar(value=True)
        self.use_zoominfo_ai_research = tk.BooleanVar(value=False)
        self.skip_history = tk.BooleanVar(value=True)
        self.resume_checkpoint = tk.BooleanVar(value=True)
        self.output_folder = tk.StringVar(value=str(OUTPUT_DIR))
        self.profile_name = tk.StringVar(value="Default")
        self.territory_zip = tk.StringVar(value="")
        self.territory_radius = tk.StringVar(value="50")
        self.research_workers = tk.IntVar(value=4)
        self.show_advanced = tk.BooleanVar(value=False)
        self.hubspot_token = tk.StringVar()
        self.hubspot_sender_email = tk.StringVar()
        self.selected_sequence = tk.StringVar()
        self.hubspot_connection_status = tk.StringVar(value="Not tested")
        self.hubspot_workspace_message = tk.StringVar(
            value="Connect HubSpot, review approved records, then sync explicitly."
        )
        self.email_intelligence_vars = {
            "public_status": tk.StringVar(value="Not evaluated"),
            "pattern": tk.StringVar(value="—"),
            "confidence": tk.StringVar(value="—"),
            "zoominfo": tk.StringVar(value="Unknown"),
            "recommendation": tk.StringVar(value="Research first"),
        }
        self.hubspot_kpi_vars = {
            "approved": tk.StringVar(value="0"),
            "ready": tk.StringVar(value="0"),
            "synced": tk.StringVar(value="0"),
            "enrolled": tk.StringVar(value="0"),
            "failed": tk.StringVar(value="0"),
        }
        self.deal_desk_filter = tk.StringVar(value="All")
        self.minimum_inbox_score = tk.IntVar(
            value=int(self.deliverability_rules.get("minimum_enrollment_score", 82))
        )
        auth = self.deliverability_rules.get("sender_authentication", {})
        self.spf_confirmed = tk.BooleanVar(value=bool(auth.get("spf_confirmed", True)))
        self.dkim_confirmed = tk.BooleanVar(value=bool(auth.get("dkim_confirmed", True)))
        self.dmarc_confirmed = tk.BooleanVar(value=bool(auth.get("dmarc_confirmed", True)))
        self.allow_low_score_override = tk.BooleanVar(value=False)
        self.delivery_outcome = tk.StringVar(value="Not Sent")
        self.use_master_dedup = tk.BooleanVar(value=True)
        self.rejected_rereview_days = tk.IntVar(value=180)
        self.qualified_rereview_days = tk.IntVar(value=0)
        self.master_status_filter = tk.StringVar(value="All")
        self.master_search = tk.StringVar(value="")
        self.discover_public_contacts = tk.BooleanVar(value=True)
        self.predict_public_emails = tk.BooleanVar(value=True)
        self.use_hubspot_csv_index = tk.BooleanVar(value=True)
        self.deep_contact_recovery = tk.BooleanVar(value=True)
        self.deep_recovery_contact_limit = tk.IntVar(value=3)
        self.block_unverified_sequence_emails = tk.BooleanVar(value=True)
        self.master_csv_path = tk.StringVar(value="")
        self.master_csv_auto_sync = tk.BooleanVar(value=True)
        self.master_csv_loaded_file = tk.StringVar(
            value="No permanent master CSV selected"
        )
        self.master_csv_company_count = tk.StringVar(value="0")
        self.master_csv_last_updated = tk.StringVar(value="Never")

        # ──────────────────────────────────────────────────────────────
        # SEARCH PROFILES — modern presentation, original 3.6 variables
        # and callbacks preserved.
        # ──────────────────────────────────────────────────────────────
        profile_frame = ttk.LabelFrame(
            settings,
            text="Discovery Profile",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        profile_frame.pack(fill="x", pady=(0, 12))

        profile_intro = ttk.Frame(profile_frame, style="Card.TFrame")
        profile_intro.grid(
            row=0,
            column=0,
            columnspan=5,
            sticky="ew",
            pady=(0, 10),
        )
        ttk.Label(
            profile_intro,
            text="Choose a saved configuration or create a new targeting profile.",
            style="Muted.TLabel",
        ).pack(side="left")
        ttk.Label(
            profile_intro,
            text="PROVEN 3.6 ENGINE",
            style="PageEyebrow.TLabel",
        ).pack(side="right")

        ttk.Label(
            profile_frame,
            text="Profile",
            style="Card.TLabel",
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 12),
        )
        self.profile_combo = ttk.Combobox(
            profile_frame,
            textvariable=self.profile_name,
            state="readonly",
            width=32,
        )
        self.profile_combo.grid(
            row=1,
            column=1,
            sticky="ew",
        )
        self.profile_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._apply_profile(),
        )
        ttk.Button(
            profile_frame,
            text="Save Profile",
            style="Blue.TButton",
            command=self._save_profile,
        ).grid(
            row=1,
            column=2,
            padx=(10, 0),
        )
        ttk.Button(
            profile_frame,
            text="Save As…",
            style="Secondary.TButton",
            command=self._save_profile_as,
        ).grid(
            row=1,
            column=3,
            padx=(8, 0),
        )
        ttk.Button(
            profile_frame,
            text="Delete",
            style="Secondary.TButton",
            command=self._delete_profile,
        ).grid(
            row=1,
            column=4,
            padx=(8, 0),
        )
        profile_frame.columnconfigure(1, weight=1)

        # Connection summary keeps credentials out of the main workflow.
        connections = ttk.LabelFrame(
            settings,
            text="Connections",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        connections.pack(fill="x", pady=(0, 12))
        connection_summary = ttk.Frame(
            connections,
            style="Card.TFrame",
        )
        connection_summary.pack(fill="x")

        zoom_status = tk.Frame(
            connection_summary,
            bg="#F4F9FD",
            highlightbackground="#CFE0ED",
            highlightthickness=1,
        )
        zoom_status.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
        )
        tk.Label(
            zoom_status,
            text="ZOOMINFO MCP",
            bg="#F4F9FD",
            fg=MUTED,
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w", padx=14, pady=(10, 2))
        self.zoominfo_connection_status = tk.StringVar(
            value=(
                "●  Configured"
                if self.client_id.get().strip()
                and self.client_secret.get().strip()
                else "○  Not configured"
            )
        )
        tk.Label(
            zoom_status,
            textvariable=self.zoominfo_connection_status,
            bg="#F4F9FD",
            fg=SUCCESS,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=14, pady=(0, 10))

        tavily_status = tk.Frame(
            connection_summary,
            bg="#F4F9FD",
            highlightbackground="#CFE0ED",
            highlightthickness=1,
        )
        tavily_status.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
        )
        tk.Label(
            tavily_status,
            text="TAVILY RESEARCH",
            bg="#F4F9FD",
            fg=MUTED,
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w", padx=14, pady=(10, 2))
        self.tavily_connection_status = tk.StringVar(
            value=(
                "●  Configured"
                if self.tavily_key.get().strip()
                else "○  Not configured"
            )
        )
        tk.Label(
            tavily_status,
            textvariable=self.tavily_connection_status,
            bg="#F4F9FD",
            fg=SUCCESS,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=14, pady=(0, 10))

        self.connections_expanded = tk.BooleanVar(value=False)
        ttk.Button(
            connection_summary,
            text="Manage Connections",
            style="Secondary.TButton",
            command=self._toggle_connection_manager,
        ).pack(side="right")
        ttk.Button(
            connection_summary,
            text="Open MCP Explorer",
            style="Blue.TButton",
            command=self._open_mcp_explorer,
        ).pack(side="right", padx=(0, 8))

        self.credentials_panel = ttk.Frame(
            connections,
            style="Card.TFrame",
        )
        self.credentials_panel.pack(
            fill="x",
            pady=(12, 0),
        )
        self._row(
            self.credentials_panel,
            0,
            "ZoomInfo MCP Client ID",
            self.client_id,
        )
        self._row(
            self.credentials_panel,
            1,
            "ZoomInfo MCP Client Secret",
            self.client_secret,
            show="•",
        )
        self._row(
            self.credentials_panel,
            2,
            "Tavily API Key",
            self.tavily_key,
            show="•",
        )
        button_col = ttk.Frame(
            self.credentials_panel,
            style="Card.TFrame",
        )
        button_col.grid(
            row=0,
            column=2,
            rowspan=3,
            padx=(12, 0),
            sticky="ns",
        )
        ttk.Button(
            button_col,
            text="Save Securely",
            style="Blue.TButton",
            command=self._save_credentials,
        ).pack(fill="x", pady=(0, 6))
        ttk.Button(
            button_col,
            text="Connect ZoomInfo",
            style="Secondary.TButton",
            command=self.connect_zoominfo,
        ).pack(fill="x")
        ttk.Button(
            button_col,
            text="Clear Authorization",
            style="Secondary.TButton",
            command=self.clear_zoominfo_authorization,
        ).pack(fill="x", pady=(6, 0))
        self.credentials_panel.pack_forget()

        # Territory card.
        territory_card = ttk.LabelFrame(
            settings,
            text="1. Territory",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        territory_card.pack(fill="x", pady=(0, 12))

        territory_modes = ttk.Frame(
            territory_card,
            style="Card.TFrame",
        )
        territory_modes.pack(fill="x", pady=(0, 10))
        ttk.Radiobutton(
            territory_modes,
            text="Selected states",
            variable=self.search_mode,
            value="states",
            command=self._update_search_mode,
        ).pack(side="left")
        ttk.Radiobutton(
            territory_modes,
            text="ZIP-radius search",
            variable=self.search_mode,
            value="radius",
            command=self._update_search_mode,
        ).pack(side="left", padx=(20, 0))

        self.state_chip_frame = tk.Frame(
            territory_card,
            bg=SURFACE,
        )
        self.state_chip_frame.pack(
            fill="x",
            pady=(0, 10),
        )
        self.state_chip_buttons = {}
        self._render_state_chips()

        territory_inputs = ttk.Frame(
            territory_card,
            style="Card.TFrame",
        )
        territory_inputs.pack(fill="x")
        ttk.Label(
            territory_inputs,
            text="Selected state codes",
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky="w")
        self.states_entry = ttk.Entry(
            territory_inputs,
            textvariable=self.states,
        )
        self.states_entry.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=(0, 14),
        )
        self.states_entry.bind(
            "<FocusOut>",
            lambda _e: self._sync_state_chips(),
        )
        self.states_entry.bind(
            "<Return>",
            lambda _e: self._sync_state_chips(),
        )
        ttk.Label(
            territory_inputs,
            text="ZIP code",
            style="Muted.TLabel",
        ).grid(row=0, column=1, sticky="w")
        self.zip_entry = ttk.Entry(
            territory_inputs,
            textvariable=self.territory_zip,
            width=14,
        )
        self.zip_entry.grid(
            row=1,
            column=1,
            sticky="w",
        )
        ttk.Label(
            territory_inputs,
            text="Radius",
            style="Muted.TLabel",
        ).grid(
            row=0,
            column=2,
            sticky="w",
            padx=(14, 0),
        )
        self.radius_combo = ttk.Combobox(
            territory_inputs,
            textvariable=self.territory_radius,
            values=["10", "25", "50", "100", "250"],
            state="readonly",
            width=10,
        )
        self.radius_combo.grid(
            row=1,
            column=2,
            sticky="w",
            padx=(14, 0),
        )
        territory_inputs.columnconfigure(0, weight=1)

        # Trades card with clickable chips backed by the original BooleanVars.
        trades_frame = ttk.LabelFrame(
            settings,
            text="2. Trades",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        trades_frame.pack(fill="x", pady=(0, 12))
        trade_header = ttk.Frame(
            trades_frame,
            style="Card.TFrame",
        )
        trade_header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            trade_header,
            text="Select the residential home-service categories to include.",
            style="Muted.TLabel",
        ).pack(side="left")
        ttk.Button(
            trade_header,
            text="Manage Trades",
            style="Secondary.TButton",
            command=self._open_trade_manager,
        ).pack(side="right")

        self.trade_checkbox_frame = tk.Frame(
            trades_frame,
            bg=SURFACE,
        )
        self.trade_checkbox_frame.pack(fill="x")
        self._render_trade_checkboxes()

        self.custom_keywords_label = ttk.Label(
            trades_frame,
            text="Custom trade keywords",
            style="Card.TLabel",
        )
        self.custom_keywords_label.pack(
            anchor="w",
            pady=(12, 4),
        )
        self.custom_keywords_entry = ttk.Entry(
            trades_frame,
            textvariable=self.custom_trade_keywords,
        )
        self.custom_keywords_entry.pack(fill="x")
        self.custom_naics_label = ttk.Label(
            trades_frame,
            text="Custom NAICS",
            style="Card.TLabel",
        )
        self.custom_naics_label.pack(
            anchor="w",
            pady=(10, 4),
        )
        self.custom_naics_entry = ttk.Entry(
            trades_frame,
            textvariable=self.naics_codes,
        )
        self.custom_naics_entry.pack(fill="x")
        self.trade_preview = tk.StringVar(value="")
        ttk.Label(
            trades_frame,
            textvariable=self.trade_preview,
            style="Muted.TLabel",
            wraplength=1050,
        ).pack(anchor="w", pady=(10, 0))

        # Company size + run goal side-by-side.
        sizing_row = ttk.Frame(
            settings,
            style="Card.TFrame",
        )
        sizing_row.pack(fill="x", pady=(0, 12))

        company_size = ttk.LabelFrame(
            sizing_row,
            text="3. Company Size",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        company_size.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 6),
        )
        size_grid = ttk.Frame(
            company_size,
            style="Card.TFrame",
        )
        size_grid.pack(fill="x")

        size_fields = [
            ("Minimum revenue", self.minimum_revenue, "$"),
            ("Maximum revenue", self.maximum_revenue, "$"),
            ("Minimum employees", self.employee_min, ""),
            ("Maximum employees", self.employee_max, ""),
        ]
        for index, (label, variable, prefix) in enumerate(size_fields):
            row = index // 2
            col = index % 2
            field = ttk.Frame(size_grid, style="Card.TFrame")
            field.grid(
                row=row,
                column=col,
                sticky="ew",
                padx=(0 if col == 0 else 8, 8 if col == 0 else 0),
                pady=(0, 10 if row == 0 else 0),
            )
            ttk.Label(
                field,
                text=label,
                style="Muted.TLabel",
            ).pack(anchor="w", pady=(0, 4))
            entry_wrap = ttk.Frame(field, style="Card.TFrame")
            entry_wrap.pack(fill="x")
            if prefix:
                ttk.Label(
                    entry_wrap,
                    text=prefix,
                    style="SectionTitle.TLabel",
                ).pack(side="left", padx=(0, 4))
            entry = ttk.Entry(
                entry_wrap,
                textvariable=variable,
                style="MetricEntry.TEntry",
            )
            entry.pack(side="left", fill="x", expand=True)
            if (
                variable is self.minimum_revenue
                or variable is self.maximum_revenue
            ):
                entry.bind(
                    "<FocusOut>",
                    lambda _e, v=variable: self._normalize_profile_number(v),
                )
                entry.bind(
                    "<Return>",
                    lambda _e, v=variable: self._normalize_profile_number(v),
                )
        size_grid.columnconfigure(0, weight=1)
        size_grid.columnconfigure(1, weight=1)

        goal_card = ttk.LabelFrame(
            sizing_row,
            text="4. Discovery Goal",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        goal_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(6, 0),
        )
        goal_grid = ttk.Frame(
            goal_card,
            style="Card.TFrame",
        )
        goal_grid.pack(fill="x")
        goal_fields = [
            ("Qualified companies", self.target_count),
            ("Candidate pool", self.candidate_limit),
            ("Contacts per company", self.contacts_per_company),
            ("Minimum fit score", self.minimum_fit_score),
        ]
        for index, (label, variable) in enumerate(goal_fields):
            row = index // 2
            col = index % 2
            field = ttk.Frame(goal_grid, style="Card.TFrame")
            field.grid(
                row=row,
                column=col,
                sticky="ew",
                padx=(0 if col == 0 else 8, 8 if col == 0 else 0),
                pady=(0, 10 if row == 0 else 0),
            )
            ttk.Label(
                field,
                text=label,
                style="Muted.TLabel",
            ).pack(anchor="w", pady=(0, 4))
            ttk.Entry(
                field,
                textvariable=variable,
                style="MetricEntry.TEntry",
            ).pack(fill="x")
        goal_grid.columnconfigure(0, weight=1)
        goal_grid.columnconfigure(1, weight=1)

        # Advanced workflow settings remain available but no longer dominate.
        advanced_card = ttk.LabelFrame(
            settings,
            text="Advanced Workflow",
            padding=14,
            style="ProfileCard.TLabelframe",
        )
        advanced_card.pack(fill="x", pady=(0, 12))
        advanced_header = ttk.Frame(
            advanced_card,
            style="Card.TFrame",
        )
        advanced_header.pack(fill="x")
        ttk.Label(
            advanced_header,
            text=(
                "Deduplication, contact recovery, credit controls, "
                "master-database policy, and output settings."
            ),
            style="Muted.TLabel",
        ).pack(side="left")
        ttk.Checkbutton(
            advanced_header,
            text="Show advanced settings",
            variable=self.show_advanced,
            command=self._update_advanced_visibility,
        ).pack(side="right")

        self.advanced_profile_panel = ttk.Frame(
            advanced_card,
            style="Card.TFrame",
        )
        self.advanced_profile_panel.pack(
            fill="x",
            pady=(12, 0),
        )

        workflow = ttk.LabelFrame(
            self.advanced_profile_panel,
            text="Workflow and Research",
            padding=10,
        )
        workflow.pack(fill="x", pady=(0, 10))
        self._row(
            workflow,
            0,
            "Parallel research workers",
            self.research_workers,
        )
        ttk.Checkbutton(
            workflow,
            text="Use ZoomInfo CRM/HubSpot flags as a secondary exclusion",
            variable=self.exclude_crm,
        ).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            workflow,
            text="Use imported HubSpot CSV index as the primary CRM source of truth",
            variable=self.use_hubspot_csv_index,
        ).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            workflow,
            text="Skip companies already reviewed by this application",
            variable=self.skip_history,
        ).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            workflow,
            text="Discover additional public decision-makers with Tavily",
            variable=self.discover_public_contacts,
        ).grid(row=4, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            workflow,
            text="Predict email only when a company pattern is supported",
            variable=self.predict_public_emails,
        ).grid(row=5, column=1, sticky="w", pady=4)
        ttk.Checkbutton(
            workflow,
            text="Run deep contact-data recovery on top-ranked contacts",
            variable=self.deep_contact_recovery,
        ).grid(row=6, column=1, sticky="w", pady=4)
        self._row(
            workflow,
            7,
            "Contacts per company for deep recovery",
            self.deep_recovery_contact_limit,
        )
        ttk.Checkbutton(
            workflow,
            text="Block predicted/unverified emails from sequence enrollment",
            variable=self.block_unverified_sequence_emails,
        ).grid(row=8, column=1, sticky="w", pady=4)

        credit_options = ttk.LabelFrame(
            workflow,
            text="Optional ZoomInfo credit-consuming features",
            padding=8,
        )
        credit_options.grid(
            row=9,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(10, 0),
        )
        ttk.Checkbutton(
            credit_options,
            text=(
                "Smart email enrichment: check ZoomInfo availability and "
                "spend only when an email is indicated"
            ),
            variable=self.smart_email_enrichment,
        ).pack(anchor="w", pady=3)
        ttk.Checkbutton(
            credit_options,
            text=(
                "Legacy bulk enrichment: enrich final contacts even when "
                "availability is unknown"
            ),
            variable=self.enrich_final_contacts,
        ).pack(anchor="w", pady=3)
        ttk.Checkbutton(
            credit_options,
            text="Use ZoomInfo AI account/contact research",
            variable=self.use_zoominfo_ai_research,
        ).pack(anchor="w", pady=3)
        ttk.Label(
            credit_options,
            text=(
                "Smart enrichment is on by default and only submits contacts "
                "whose raw ZoomInfo results indicate email availability. "
                "Legacy bulk enrichment remains off because it may spend "
                "credits without a confirmed email gain."
            ),
            style="Card.TLabel",
            foreground=WARNING,
            wraplength=900,
        ).pack(anchor="w", pady=(4, 0))

        master_policy = ttk.LabelFrame(
            self.advanced_profile_panel,
            text="Master Prospect Database Policy",
            padding=10,
        )
        master_policy.pack(fill="x", pady=(0, 10))
        ttk.Checkbutton(
            master_policy,
            text="Use the master database to prevent repeat qualification",
            variable=self.use_master_dedup,
        ).grid(row=0, column=1, sticky="w", pady=4)
        self._row(
            master_policy,
            1,
            "Re-review rejected companies after days",
            self.rejected_rereview_days,
        )
        self._row(
            master_policy,
            2,
            "Re-review qualified companies after days",
            self.qualified_rereview_days,
        )
        ttk.Label(
            master_policy,
            text=(
                "Use 0 for qualified companies to skip them permanently. "
                "Approved, HubSpot-synced, sequenced, meeting, customer, "
                "and lost records remain protected."
            ),
            style="Muted.TLabel",
            wraplength=920,
        ).grid(
            row=3,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(8, 0),
        )

        destination = ttk.LabelFrame(
            self.advanced_profile_panel,
            text="Output and Application",
            padding=10,
        )
        destination.pack(fill="x")
        self._row(
            destination,
            0,
            "Output folder",
            self.output_folder,
        )
        ttk.Button(
            destination,
            text="Browse",
            style="Secondary.TButton",
            command=self._browse_output_folder,
        ).grid(row=0, column=2, padx=(8, 0))
        ttk.Checkbutton(
            destination,
            text="Resume the previous interrupted run when a checkpoint exists",
            variable=self.resume_checkpoint,
        ).grid(row=1, column=1, sticky="w", pady=4)
        app_buttons = ttk.Frame(
            destination,
            style="Card.TFrame",
        )
        app_buttons.grid(
            row=2,
            column=1,
            sticky="w",
            pady=(8, 0),
        )
        ttk.Button(
            app_buttons,
            text="Create Desktop Shortcut",
            style="Secondary.TButton",
            command=self._create_shortcut,
        ).pack(side="left")
        ttk.Button(
            app_buttons,
            text="Check for Updates",
            style="Secondary.TButton",
            command=self._check_updates,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            app_buttons,
            text="Open Output Folder",
            style="Secondary.TButton",
            command=self._open_output_folder,
        ).pack(side="left", padx=(8, 0))

        # Keep advanced content collapsed by default.
        self.advanced_profile_panel.pack_forget()
        self._update_advanced_visibility()

        # Primary action is intentionally unmistakable.
        launch_card = tk.Frame(
            settings,
            bg=NAVY,
            highlightthickness=0,
        )
        launch_card.pack(fill="x", pady=(0, 6))
        launch_copy = tk.Frame(
            launch_card,
            bg=NAVY,
        )
        launch_copy.pack(
            side="left",
            fill="both",
            expand=True,
            padx=18,
            pady=14,
        )
        tk.Label(
            launch_copy,
            text="READY TO BUILD THE NEXT QUALIFIED LIST?",
            bg=NAVY,
            fg="#78B7F4",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            launch_copy,
            text="Run the proven discovery workflow with this profile.",
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 14),
        ).pack(anchor="w", pady=(2, 0))
        launch_buttons = tk.Frame(
            launch_card,
            bg=NAVY,
        )
        launch_buttons.pack(
            side="right",
            padx=18,
            pady=14,
        )
        ttk.Button(
            launch_buttons,
            text="Save Settings",
            style="Secondary.TButton",
            command=self._save_settings,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            launch_buttons,
            text="START DISCOVERY  →",
            style="Hero.TButton",
            command=self.start,
        ).pack(side="left")

        self._refresh_trade_preview()
        self._sync_state_chips()
        self._update_search_mode()

        dashboard_header = tk.Frame(
            run, bg=NAVY, padx=20, pady=16,
            highlightthickness=0,
        )
        dashboard_header.pack(fill="x", pady=(0, 12))
        tk.Label(
            dashboard_header,
            text="LIVE PROSPECT INTELLIGENCE",
            bg=NAVY,
            fg="#75B6F5",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            dashboard_header,
            text="Run Dashboard",
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 20),
        ).pack(anchor="w", pady=(2, 2))
        tk.Label(
            dashboard_header,
            text=(
                "See exactly where every candidate is in the proven workflow—"
                "from ZoomInfo discovery through research, qualification, "
                "contact recovery, and review."
            ),
            bg=NAVY,
            fg="#CFE3F6",
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        self.metric_vars = {
            "qualified": tk.StringVar(value="0"),
            "contacts": tk.StringVar(value="0"),
            "reviewed": tk.StringVar(value="0"),
            "elapsed": tk.StringVar(value="00:00"),
            "eta": tk.StringVar(value="—"),
        }
        self.funnel_vars = {
            "zoominfo": tk.StringVar(value="0"),
            "reviewed": tk.StringVar(value="0"),
            "researched": tk.StringVar(value="0"),
            "rejected": tk.StringVar(value="0"),
            "qualified": tk.StringVar(value="0"),
            "contacts": tk.StringVar(value="0"),
        }
        self.funnel_counts = {
            "zoominfo": 0,
            "reviewed": 0,
            "researched": 0,
            "rejected": 0,
            "qualified": 0,
            "contacts": 0,
        }
        self.preview_vars = {
            "company": tk.StringVar(value="Waiting for discovery to begin"),
            "phase": tk.StringVar(value="READY"),
            "location": tk.StringVar(value="—"),
            "size": tk.StringVar(value="—"),
            "score": tk.StringVar(value="—"),
            "residential": tk.StringVar(value="—"),
            "growth": tk.StringVar(value="—"),
            "contact": tk.StringVar(value="No contact ranked yet"),
            "detail": tk.StringVar(
                value="Start Discovery to see live company intelligence."
            ),
        }

        # Primary KPIs.
        dashboard = tk.Frame(run, bg=SURFACE)
        dashboard.pack(fill="x", pady=(0, 10))
        metric_specs = [
            ("Qualified", "qualified", "Ready for review"),
            ("Contacts", "contacts", "Ranked decision-makers"),
            ("Reviewed", "reviewed", "Unique candidates evaluated"),
            ("Elapsed", "elapsed", "Current run time"),
            ("Remaining", "eta", "Estimated from live pace"),
        ]
        for col, (label, key, helper) in enumerate(metric_specs):
            card = tk.Frame(
                dashboard,
                bg=SURFACE,
                padx=16,
                pady=13,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=col,
                sticky="nsew",
                padx=(0 if col == 0 else 5, 0),
            )
            tk.Frame(card, bg=ACCENT, height=4).pack(
                fill="x", pady=(0, 9)
            )
            tk.Label(
                card,
                textvariable=self.metric_vars[key],
                bg=SURFACE,
                fg=NAVY,
                font=("Segoe UI Semibold", 23),
            ).pack(anchor="w")
            tk.Label(
                card,
                text=label,
                bg=SURFACE,
                fg=TEXT,
                font=("Segoe UI Semibold", 9),
            ).pack(anchor="w", pady=(2, 0))
            tk.Label(
                card,
                text=helper,
                bg=SURFACE,
                fg=MUTED,
                font=("Segoe UI", 8),
            ).pack(anchor="w", pady=(2, 0))
            dashboard.columnconfigure(col, weight=1)

        controls_card = tk.Frame(
            run,
            bg=SURFACE,
            padx=14,
            pady=11,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        controls_card.pack(fill="x", pady=(0, 10))
        controls = ttk.Frame(controls_card, style="Card.TFrame")
        controls.pack(fill="x")
        self.run_button = ttk.Button(
            controls,
            text="Start Discovery",
            style="Primary.TButton",
            command=self.start,
        )
        self.run_button.pack(side="left")
        self.stop_button = ttk.Button(
            controls,
            text="Stop",
            style="Secondary.TButton",
            command=self.stop,
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=(8, 0))
        self.progress = ttk.Progressbar(controls, mode="indeterminate")
        self.progress.pack(
            side="left",
            fill="x",
            expand=True,
            padx=14,
        )
        self.status = tk.StringVar(value="Ready")
        tk.Label(
            controls_card,
            text="●",
            bg=SURFACE,
            fg="#19A66A",
            font=("Segoe UI", 11),
        ).pack(side="left", pady=(8, 0))
        tk.Label(
            controls_card,
            textvariable=self.status,
            bg=SURFACE,
            fg=NAVY,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left", padx=(7, 0), pady=(8, 0))

        # Funnel + current-company intelligence.
        intelligence_row = tk.Frame(run, bg=SURFACE)
        intelligence_row.pack(fill="x", pady=(0, 10))

        funnel_card = tk.Frame(
            intelligence_row,
            bg=SURFACE,
            padx=14,
            pady=12,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        funnel_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 5),
        )
        tk.Label(
            funnel_card,
            text="LIVE DISCOVERY FUNNEL",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            funnel_card,
            text="How the candidate pool is moving through Compass",
            bg=SURFACE,
            fg=NAVY,
            font=("Segoe UI Semibold", 12),
        ).pack(anchor="w", pady=(2, 10))

        funnel_stages = [
            ("ZoomInfo found", "zoominfo"),
            ("Reviewed", "reviewed"),
            ("Tavily researched", "researched"),
            ("Rejected / filtered", "rejected"),
            ("Qualified", "qualified"),
            ("Contacts ranked", "contacts"),
        ]
        self.funnel_stage_frames = {}
        for index, (label, key) in enumerate(funnel_stages):
            stage = tk.Frame(
                funnel_card,
                bg="#F4F8FC",
                padx=10,
                pady=7,
                highlightthickness=1,
                highlightbackground="#D8E4EF",
            )
            stage.pack(fill="x", pady=(0, 5))
            tk.Label(
                stage,
                text=str(index + 1),
                bg=ACCENT if key in {"qualified", "contacts"} else "#DDEAF5",
                fg=WHITE if key in {"qualified", "contacts"} else NAVY,
                width=3,
                font=("Segoe UI Semibold", 9),
            ).pack(side="left")
            tk.Label(
                stage,
                text=label,
                bg="#F4F8FC",
                fg=TEXT,
                font=("Segoe UI Semibold", 9),
            ).pack(side="left", padx=(9, 0))
            tk.Label(
                stage,
                textvariable=self.funnel_vars[key],
                bg="#F4F8FC",
                fg=ACCENT if key in {"qualified", "contacts"} else NAVY,
                font=("Segoe UI Semibold", 13),
            ).pack(side="right")
            self.funnel_stage_frames[key] = stage

        preview_card = tk.Frame(
            intelligence_row,
            bg=NAVY,
            padx=16,
            pady=13,
            highlightthickness=0,
        )
        preview_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(5, 0),
        )
        top_line = tk.Frame(preview_card, bg=NAVY)
        top_line.pack(fill="x")
        tk.Label(
            top_line,
            text="CURRENT PROSPECT",
            bg=NAVY,
            fg="#75B6F5",
            font=("Segoe UI Semibold", 8),
        ).pack(side="left")
        tk.Label(
            top_line,
            textvariable=self.preview_vars["phase"],
            bg="#0F568F",
            fg=WHITE,
            padx=9,
            pady=3,
            font=("Segoe UI Semibold", 8),
        ).pack(side="right")
        tk.Label(
            preview_card,
            textvariable=self.preview_vars["company"],
            bg=NAVY,
            fg=WHITE,
            anchor="w",
            font=("Segoe UI Semibold", 16),
        ).pack(fill="x", pady=(7, 2))
        tk.Label(
            preview_card,
            textvariable=self.preview_vars["detail"],
            bg=NAVY,
            fg="#CFE3F6",
            anchor="w",
            justify="left",
            wraplength=620,
            font=("Segoe UI", 9),
        ).pack(fill="x", pady=(0, 11))

        preview_grid = tk.Frame(preview_card, bg=NAVY)
        preview_grid.pack(fill="x")
        preview_specs = [
            ("LOCATION", "location"),
            ("COMPANY SIZE", "size"),
            ("FIT SCORE", "score"),
            ("RESIDENTIAL", "residential"),
            ("GROWTH", "growth"),
        ]
        for col, (label, key) in enumerate(preview_specs):
            box = tk.Frame(
                preview_grid,
                bg="#0D3558",
                padx=9,
                pady=8,
            )
            box.grid(
                row=0,
                column=col,
                sticky="nsew",
                padx=(0 if col == 0 else 4, 0),
            )
            tk.Label(
                box,
                text=label,
                bg="#0D3558",
                fg="#7FAACD",
                font=("Segoe UI Semibold", 7),
            ).pack(anchor="w")
            tk.Label(
                box,
                textvariable=self.preview_vars[key],
                bg="#0D3558",
                fg=WHITE,
                font=("Segoe UI Semibold", 10),
                wraplength=120,
                justify="left",
            ).pack(anchor="w", pady=(3, 0))
            preview_grid.columnconfigure(col, weight=1)

        contact_box = tk.Frame(
            preview_card,
            bg="#0D3558",
            padx=10,
            pady=8,
        )
        contact_box.pack(fill="x", pady=(9, 0))
        tk.Label(
            contact_box,
            text="TOP CONTACT / WHY",
            bg="#0D3558",
            fg="#7FAACD",
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            contact_box,
            textvariable=self.preview_vars["contact"],
            bg="#0D3558",
            fg=WHITE,
            anchor="w",
            justify="left",
            wraplength=650,
            font=("Segoe UI", 9),
        ).pack(fill="x", pady=(3, 0))

        log_frame = ttk.LabelFrame(
            run,
            text="Live Activity",
            padding=10,
        )
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(
            log_frame,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
            bg="#071B2E",
            fg="#DCEBFA",
            insertbackground=WHITE,
            relief="flat",
            highlightthickness=1,
            highlightbackground="#173E61",
            padx=14,
            pady=12,
            height=12,
        )
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # Deal Desk: human approval gate before HubSpot or sequence enrollment.
        desk_header = tk.Frame(
            deal_desk_tab, bg=NAVY, padx=20, pady=15
        )
        desk_header.pack(fill="x", pady=(0, 12))
        title_area = tk.Frame(desk_header, bg=NAVY)
        title_area.pack(side="left", fill="x", expand=True)
        tk.Label(
            title_area, text="REVIEW & APPROVAL",
            bg=NAVY, fg="#75B6F5",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            title_area, text="Deal Desk",
            bg=NAVY, fg=WHITE,
            font=("Segoe UI Semibold", 19),
        ).pack(anchor="w", pady=(2, 1))
        tk.Label(
            title_area,
            text="Review research, contacts, outreach, and inbox readiness before any HubSpot action.",
            bg=NAVY, fg="#CFE3F6",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        filter_wrap = tk.Frame(desk_header, bg=NAVY)
        filter_wrap.pack(side="right", padx=(18, 0))
        tk.Label(
            filter_wrap, text="SHOW", bg=NAVY, fg="#9FC3E2",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="e", pady=(0, 4))
        desk_filter = ttk.Combobox(
            filter_wrap,
            textvariable=self.deal_desk_filter,
            state="readonly",
            values=[
                "All", "Pending Review", "Needs Verification",
                "Approved", "Rejected", "Synced", "Enrolled",
            ],
            width=18,
        )
        desk_filter.pack(side="left")
        desk_filter.bind(
            "<<ComboboxSelected>>",
            lambda _event: self._refresh_deal_desk(),
        )

        self.deal_kpi_vars = {
            "pending": tk.StringVar(value="0"),
            "verification": tk.StringVar(value="0"),
            "approved": tk.StringVar(value="0"),
            "synced": tk.StringVar(value="0"),
            "enrolled": tk.StringVar(value="0"),
        }

        desk_kpis = tk.Frame(
            deal_desk_tab,
            bg=SURFACE,
        )
        desk_kpis.pack(fill="x", pady=(0, 10))
        desk_kpi_specs = [
            ("Pending Review", "pending", "#EAF3FB", NAVY),
            ("Needs Verification", "verification", "#FFF4DB", WARNING),
            ("Approved", "approved", "#E9F7F0", SUCCESS),
            ("Synced", "synced", "#EAF3FB", ACCENT),
            ("Enrolled", "enrolled", "#F1ECFB", "#6545A4"),
        ]
        for column, (label, key, background, foreground) in enumerate(
            desk_kpi_specs
        ):
            card = tk.Frame(
                desk_kpis,
                bg=background,
                padx=11,
                pady=6,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 5, 0),
            )
            tk.Label(
                card,
                textvariable=self.deal_kpi_vars[key],
                bg=background,
                fg=foreground,
                font=("Segoe UI Semibold", 15),
            ).pack(anchor="w")
            tk.Label(
                card,
                text=label,
                bg=background,
                fg=TEXT,
                font=("Segoe UI Semibold", 8),
            ).pack(anchor="w", pady=(2, 0))
            desk_kpis.columnconfigure(column, weight=1)

        desk_pane = ttk.Panedwindow(
            deal_desk_tab, orient="horizontal"
        )
        desk_pane.pack(fill="both", expand=True)

        queue_frame = ttk.LabelFrame(
            desk_pane, text="Prospects and Contacts", padding=8
        )
        review_frame = ttk.LabelFrame(
            desk_pane, text="Review and Edit", padding=10
        )
        desk_pane.add(queue_frame, weight=3)
        desk_pane.add(review_frame, weight=4)
        self.after(
            250,
            lambda: self._set_initial_deal_desk_split(desk_pane),
        )

        queue_toolbar = tk.Frame(
            queue_frame,
            bg=SURFACE,
        )
        queue_toolbar.pack(fill="x", pady=(0, 6))

        tk.Label(
            queue_toolbar,
            text="Review Queue",
            bg=SURFACE,
            fg=NAVY,
            font=("Segoe UI Semibold", 10),
        ).pack(side="left")

        tk.Label(
            queue_toolbar,
            text="Search globally or filter by workflow status.",
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 8),
        ).pack(side="left", padx=(8, 0))

        queue_columns = (
            "order", "status", "company", "contact", "strategy",
            "confidence", "inbox", "hubspot", "enrollment",
        )
        queue_table_frame = ttk.Frame(
            queue_frame,
            style="Card.TFrame",
        )
        queue_table_frame.pack(fill="both", expand=True)

        self.deal_tree = ttk.Treeview(
            queue_table_frame,
            columns=queue_columns,
            show="headings",
        )
        queue_headings = {
            "order": "Outreach Order",
            "status": "Review Status",
            "company": "Company",
            "contact": "Contact",
            "strategy": "Strategy",
            "confidence": "Outreach",
            "inbox": "Inbox",
            "hubspot": "HubSpot",
            "enrollment": "Sequence",
        }
        widths = {
            "order": 110, "status": 100, "company": 170, "contact": 165,
            "strategy": 170, "confidence": 70, "inbox": 70,
            "hubspot": 90, "enrollment": 95,
        }
        for name in queue_columns:
            self.deal_tree.heading(name, text=queue_headings[name])
            self.deal_tree.column(name, width=widths[name], stretch=True)
        configure_scrollable_tree(
            queue_table_frame,
            self.deal_tree,
            horizontal=True,
            vertical=True,
            enable_mousewheel=True,
        )
        self.deal_tree.bind(
            "<<TreeviewSelect>>",
            lambda _event: self._load_selected_queue_item(),
        )
        self.deal_tree.tag_configure(
            "pending",
            background="#F8FBFE",
            foreground=TEXT,
        )
        self.deal_tree.tag_configure(
            "verification",
            background="#FFF7E5",
            foreground="#7A4A00",
        )
        self.deal_tree.tag_configure(
            "approved",
            background="#ECF8F2",
            foreground="#12633E",
        )
        self.deal_tree.tag_configure(
            "rejected",
            background="#FDEEEE",
            foreground="#8A221B",
        )
        self.deal_tree.tag_configure(
            "synced",
            background="#EDF5FC",
            foreground="#064E8B",
        )
        self.deal_tree.tag_configure(
            "enrolled",
            background="#F2EEFA",
            foreground="#55348C",
        )

        self.review_queue_id = tk.StringVar()
        self.review_company = tk.StringVar()
        self.review_contact = tk.StringVar()
        self.review_strategy = tk.StringVar()
        self.review_confidence = tk.StringVar()
        self.review_outreach_order = tk.StringVar(
            value="Select a contact to see outreach priority"
        )
        self.review_contact_source = tk.StringVar(value="")
        self.review_subject = tk.StringVar()
        self.review_notes = tk.StringVar()
        self.review_metric_vars = {
            "decision": tk.StringVar(value="—"),
            "email": tk.StringVar(value="—"),
            "phone": tk.StringVar(value="—"),
            "inbox": tk.StringVar(value="—"),
            "outreach": tk.StringVar(value="—"),
        }
        self.review_company_detail = tk.StringVar(
            value="Select a company and contact from the review queue."
        )
        self.review_recommendation = tk.StringVar(
            value="Contact recommendation will appear here."
        )
        self.review_email_detail = tk.StringVar(
            value="Email quality and verification status"
        )
        self.review_phone_detail = tk.StringVar(
            value="Phone quality and source status"
        )

        workflow_bar = tk.Frame(
            review_frame,
            bg="#EAF3FB",
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=10,
            pady=8,
        )
        workflow_bar.pack(fill="x", pady=(0, 8))

        tk.Label(
            workflow_bar,
            text="COMPANY WORKFLOW",
            bg="#EAF3FB",
            fg=NAVY,
            font=("Segoe UI Semibold", 8),
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            workflow_bar,
            text="Approve Company",
            style="Primary.TButton",
            command=self._approve_selected_company,
        ).pack(side="left")

        ttk.Button(
            workflow_bar,
            text="Needs Verification",
            style="Secondary.TButton",
            command=self._mark_selected_company_for_verification,
        ).pack(side="left", padx=(8, 0))

        ttk.Button(
            workflow_bar,
            text="Reject Company",
            style="Secondary.TButton",
            command=self._reject_selected_company,
        ).pack(side="left", padx=(8, 0))

        ttk.Button(
            workflow_bar,
            text="Preview Email",
            style="Secondary.TButton",
            command=self._show_email_review_tab,
        ).pack(side="right")

        self.expand_tabs_button = ttk.Button(
            workflow_bar,
            text="Expand Tabs",
            style="Secondary.TButton",
            command=self._toggle_deal_desk_tabs,
        )
        self.expand_tabs_button.pack(side="right", padx=(0, 8))

        ttk.Button(
            workflow_bar,
            text="Sync Company to HubSpot",
            style="Primary.TButton",
            command=self._sync_selected_company_to_hubspot,
        ).pack(side="right", padx=(0, 8))

        review_vertical_pane = ttk.Panedwindow(
            review_frame,
            orient="vertical",
        )
        review_vertical_pane.pack(
            fill="both",
            expand=True,
            pady=(6, 0),
        )

        review_summary_pane = ttk.Frame(
            review_vertical_pane,
            style="Card.TFrame",
        )
        review_tabs_pane = ttk.Frame(
            review_vertical_pane,
            style="Card.TFrame",
        )
        review_vertical_pane.add(review_summary_pane, weight=2)
        review_vertical_pane.add(review_tabs_pane, weight=5)

        self.review_vertical_pane = review_vertical_pane
        self.review_summary_pane = review_summary_pane
        self.review_tabs_pane = review_tabs_pane
        self.review_tabs_expanded = False
        self.after(
            300,
            self._set_initial_review_vertical_split,
        )

        summary_card = tk.Frame(
            review_summary_pane,
            bg=SURFACE,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=10,
            pady=8,
        )
        summary_card.pack(fill="x")

        summary_top = tk.Frame(summary_card, bg=SURFACE)
        summary_top.pack(fill="x")

        priority_badge = tk.Label(
            summary_top,
            textvariable=self.review_outreach_order,
            bg=ACCENT,
            fg=WHITE,
            font=("Segoe UI Semibold", 9),
            padx=10,
            pady=6,
        )
        priority_badge.pack(side="left")

        identity_block = tk.Frame(summary_top, bg=SURFACE)
        identity_block.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(10, 8),
        )
        tk.Label(
            identity_block,
            textvariable=self.review_company,
            bg=SURFACE,
            fg=NAVY,
            font=("Segoe UI Semibold", 14),
        ).pack(anchor="w")
        tk.Label(
            identity_block,
            textvariable=self.review_contact,
            bg=SURFACE,
            fg=TEXT,
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w", pady=(1, 0))
        tk.Label(
            identity_block,
            textvariable=self.review_strategy,
            bg=SURFACE,
            fg=ACCENT,
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(1, 0))
        tk.Label(
            identity_block,
            textvariable=self.review_company_detail,
            bg=SURFACE,
            fg=MUTED,
            font=("Segoe UI", 8),
            justify="left",
            wraplength=560,
        ).pack(anchor="w", pady=(1, 0))

        outreach_badge = tk.Frame(
            summary_top,
            bg=NAVY,
            padx=10,
            pady=6,
        )
        outreach_badge.pack(side="right")
        tk.Label(
            outreach_badge,
            text="OUTREACH",
            bg=NAVY,
            fg="#87B7DD",
            font=("Segoe UI Semibold", 6),
        ).pack()
        tk.Label(
            outreach_badge,
            textvariable=self.review_metric_vars["outreach"],
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 17),
        ).pack()

        compact_scores = tk.Frame(summary_card, bg=SURFACE)
        compact_scores.pack(fill="x", pady=(7, 0))
        score_specs = [
            ("Decision", "decision"),
            ("Email", "email"),
            ("Phone", "phone"),
            ("Inbox", "inbox"),
        ]
        for column, (label, key) in enumerate(score_specs):
            score = tk.Frame(
                compact_scores,
                bg="#F4F8FC",
                padx=8,
                pady=5,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            score.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 4, 0),
            )
            tk.Label(
                score,
                text=label.upper(),
                bg="#F4F8FC",
                fg=MUTED,
                font=("Segoe UI Semibold", 6),
            ).pack(anchor="w")
            tk.Label(
                score,
                textvariable=self.review_metric_vars[key],
                bg="#F4F8FC",
                fg=ACCENT if key != "inbox" else WARNING,
                font=("Segoe UI Semibold", 12),
            ).pack(anchor="w")
            compact_scores.columnconfigure(column, weight=1)

        why_row = tk.Frame(
            summary_card,
            bg="#F8FBFE",
            padx=8,
            pady=5,
            highlightthickness=1,
            highlightbackground="#DFE8F0",
        )
        why_row.pack(fill="x", pady=(6, 0))
        tk.Label(
            why_row,
            text="WHY THIS CONTACT",
            bg="#F8FBFE",
            fg=MUTED,
            font=("Segoe UI Semibold", 6),
        ).pack(anchor="w")
        why_contact_value = tk.Label(
            why_row,
            textvariable=self.review_recommendation,
            bg="#F8FBFE",
            fg=TEXT,
            font=("Segoe UI", 8),
            justify="left",
            anchor="w",
            wraplength=500,
        )
        why_contact_value.pack(fill="x", anchor="w", pady=(2, 0))

        def resize_why_contact(event):
            why_contact_value.configure(
                wraplength=max(180, event.width - 18)
            )

        why_row.bind("<Configure>", resize_why_contact, add="+")

        review_notebook = ttk.Notebook(review_tabs_pane)
        review_notebook.pack(
            fill="both",
            expand=True,
        )
        email_page = ttk.Frame(
            review_notebook,
            style="Card.TFrame",
        )
        company_page = ttk.Frame(
            review_notebook,
            style="Card.TFrame",
        )
        intelligence_page = ttk.Frame(
            review_notebook,
            style="Card.TFrame",
        )
        acquisition_page = ttk.Frame(
            review_notebook,
            style="Card.TFrame",
        )

        email_scroller = VerticalScrolledFrame(
            email_page,
            background=SURFACE,
            padding=8,
        )
        company_scroller = VerticalScrolledFrame(
            company_page,
            background=SURFACE,
            padding=8,
        )
        intelligence_scroller = VerticalScrolledFrame(
            intelligence_page,
            background=SURFACE,
            padding=8,
        )
        acquisition_scroller = VerticalScrolledFrame(
            acquisition_page,
            background=SURFACE,
            padding=8,
        )
        for scroller in (
            email_scroller,
            company_scroller,
            intelligence_scroller,
            acquisition_scroller,
        ):
            scroller.pack(fill="both", expand=True)

        email_tab = email_scroller.content
        company_tab = company_scroller.content
        intelligence_tab = intelligence_scroller.content
        acquisition_tab = acquisition_scroller.content

        review_notebook.add(email_page, text="1. Email Review")
        review_notebook.add(
            company_page, text="2. Company Intelligence"
        )
        review_notebook.add(
            intelligence_page, text="3. Contact Intelligence"
        )
        review_notebook.add(
            acquisition_page, text="4. Contact Acquisition Report"
        )
        self.review_notebook = review_notebook
        self.email_tab = email_page
        self.company_tab = company_page
        self.intelligence_tab = intelligence_page
        self.acquisition_tab = acquisition_page
        self.review_tab_scrollers = {
            "email": email_scroller,
            "company": company_scroller,
            "contact": intelligence_scroller,
            "acquisition": acquisition_scroller,
        }
        def reset_selected_review_tab(_event=None):
            selected = review_notebook.select()
            scroller_by_page = {
                str(email_page): email_scroller,
                str(company_page): company_scroller,
                str(intelligence_page): intelligence_scroller,
                str(acquisition_page): acquisition_scroller,
            }
            scroller = scroller_by_page.get(selected)
            if scroller is not None:
                self.after_idle(scroller.scroll_to_top)

        review_notebook.bind(
            "<<NotebookTabChanged>>",
            reset_selected_review_tab,
            add="+",
        )

        email_tab.columnconfigure(0, weight=1)

        email_header = ttk.Frame(
            email_tab,
            style="Card.TFrame",
        )
        email_header.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 6),
        )
        ttk.Label(
            email_header,
            text="Approved Email",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        ttk.Label(
            email_header,
            text="Review the complete message before approval or sync.",
            style="Muted.TLabel",
            wraplength=620,
            justify="left",
        ).pack(fill="x", anchor="w", pady=(2, 0))

        subject_frame = ttk.Frame(
            email_tab,
            style="Card.TFrame",
        )
        subject_frame.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 6),
        )
        subject_frame.columnconfigure(0, weight=1)
        ttk.Label(
            subject_frame,
            text="Approved Subject",
            style="Card.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 3))
        ttk.Entry(
            subject_frame,
            textvariable=self.review_subject,
        ).grid(row=1, column=0, sticky="ew")

        ttk.Label(
            email_tab,
            text="Approved Email Body",
            style="Card.TLabel",
        ).grid(
            row=2,
            column=0,
            sticky="w",
            pady=(0, 3),
        )

        email_body_frame = ttk.Frame(
            email_tab,
            style="Card.TFrame",
        )
        email_body_frame.grid(
            row=3,
            column=0,
            sticky="ew",
        )
        self.review_email_body = tk.Text(
            email_body_frame,
            height=24,
            width=90,
            wrap="word",
            bg=WHITE,
            fg=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            font=("Segoe UI", 10),
            undo=True,
            padx=14,
            pady=12,
        )
        configure_scrollable_text(
            email_body_frame,
            self.review_email_body,
            horizontal=False,
            vertical=True,
            enable_mousewheel=True,
        )

        notes_frame = ttk.Frame(
            email_tab,
            style="Card.TFrame",
        )
        notes_frame.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(7, 0),
        )
        notes_frame.columnconfigure(0, weight=1)
        ttk.Label(
            notes_frame,
            text="Reviewer Notes",
            style="Card.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 3))
        ttk.Entry(
            notes_frame,
            textvariable=self.review_notes,
        ).grid(row=1, column=0, sticky="ew")

        company_header = ttk.Frame(
            company_tab,
            style="Card.TFrame",
        )
        company_header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            company_header,
            text="Company Intelligence",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        company_header_description = ttk.Label(
            company_header,
            text=(
                "Qualification evidence, residential fit, company size, "
                "growth signals, technology, and the Darwill opportunity."
            ),
            style="Muted.TLabel",
            wraplength=600,
            justify="left",
        )
        company_header_description.pack(fill="x", anchor="w", pady=(2, 0))
        company_header.bind(
            "<Configure>",
            lambda event: company_header_description.configure(
                wraplength=max(220, event.width - 12)
            ),
            add="+",
        )

        self.company_intelligence_vars.update({
            "qualification": tk.StringVar(value="—"),
            "residential": tk.StringVar(value="—"),
            "size": tk.StringVar(value="—"),
            "opportunity": tk.StringVar(value="—"),
            "location": tk.StringVar(value="—"),
            "website": tk.StringVar(value="—"),
            "growth": tk.StringVar(value="No stored growth signal"),
            "technology": tk.StringVar(value="No stored technology signal"),
            "service_area": tk.StringVar(value="No stored service-area evidence"),
            "darwill_angle": tk.StringVar(
                value="Select a company to calculate the recommended Darwill angle."
            ),
        })

        executive_summary_card = tk.Frame(
            company_tab,
            bg="#F4F8FC",
            padx=12,
            pady=10,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        executive_summary_card.pack(fill="x", pady=(0, 10))

        tk.Label(
            executive_summary_card,
            text="EXECUTIVE SUMMARY",
            bg="#F4F8FC",
            fg=ACCENT,
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")

        executive_summary_value = tk.Label(
            executive_summary_card,
            textvariable=self.company_intelligence_vars[
                "executive_summary"
            ],
            bg="#F4F8FC",
            fg=TEXT,
            font=("Segoe UI", 10),
            justify="left",
            anchor="nw",
            wraplength=580,
        )
        executive_summary_value.pack(
            fill="x",
            anchor="w",
            pady=(4, 0),
        )

        executive_summary_card.bind(
            "<Configure>",
            lambda event: executive_summary_value.configure(
                wraplength=max(240, event.width - 24)
            ),
            add="+",
        )

        company_score_row = tk.Frame(company_tab, bg=SURFACE)
        company_score_row.pack(fill="x", pady=(0, 8))
        company_score_specs = [
            ("Qualification Evidence", "qualification", "#EDF5FC", ACCENT),
            ("Residential Confidence", "residential", "#ECF8F2", SUCCESS),
            ("Company Size", "size", "#F4F0FB", "#6545A4"),
            ("Darwill Opportunity", "opportunity", "#FFF4DB", WARNING),
        ]
        for column, (label, key, background, foreground) in enumerate(
            company_score_specs
        ):
            card = tk.Frame(
                company_score_row,
                bg=background,
                padx=10,
                pady=8,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 5, 0),
            )
            score_heading = tk.Label(
                card,
                text=label.upper(),
                bg=background,
                fg=MUTED,
                font=("Segoe UI Semibold", 7),
                justify="left",
                anchor="w",
                wraplength=145,
            )
            score_heading.pack(fill="x", anchor="w")
            score_value = tk.Label(
                card,
                textvariable=self.company_intelligence_vars[key],
                bg=background,
                fg=foreground,
                font=("Segoe UI Semibold", 15),
                wraplength=145,
                justify="left",
                anchor="w",
            )
            score_value.pack(fill="x", anchor="w", pady=(2, 0))

            def resize_company_score(
                event,
                heading=score_heading,
                value=score_value,
            ):
                available = max(90, event.width - 20)
                heading.configure(wraplength=available)
                value.configure(wraplength=available)

            card.bind("<Configure>", resize_company_score, add="+")
            company_score_row.columnconfigure(column, weight=1)

        company_facts = tk.Frame(company_tab, bg=SURFACE)
        company_facts.pack(fill="x", pady=(0, 8))
        fact_specs = [
            ("LOCATION", "location"),
            ("WEBSITE", "website"),
            ("SERVICE AREA", "service_area"),
        ]
        for column, (label, key) in enumerate(fact_specs):
            card = tk.Frame(
                company_facts,
                bg="#F4F8FC",
                padx=10,
                pady=8,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 5, 0),
            )
            tk.Label(
                card,
                text=label,
                bg="#F4F8FC",
                fg=MUTED,
                font=("Segoe UI Semibold", 7),
            ).pack(anchor="w")
            fact_value = tk.Label(
                card,
                textvariable=self.company_intelligence_vars[key],
                bg="#F4F8FC",
                fg=TEXT,
                font=("Segoe UI", 9),
                wraplength=190,
                justify="left",
                anchor="w",
            )
            fact_value.pack(fill="x", anchor="w", pady=(3, 0))
            card.bind(
                "<Configure>",
                lambda event, widget=fact_value: widget.configure(
                    wraplength=max(100, event.width - 20)
                ),
                add="+",
            )
            company_facts.columnconfigure(column, weight=1)

        signal_row = tk.Frame(company_tab, bg=SURFACE)
        signal_row.pack(fill="x", pady=(0, 8))
        growth_card = tk.Frame(
            signal_row,
            bg="#ECF8F2",
            padx=11,
            pady=8,
            highlightthickness=1,
            highlightbackground="#CBE8D9",
        )
        growth_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 4),
        )
        tk.Label(
            growth_card,
            text="GROWTH & MARKETING SIGNALS",
            bg="#ECF8F2",
            fg=SUCCESS,
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            growth_card,
            textvariable=self.company_intelligence_vars["growth"],
            bg="#ECF8F2",
            fg=TEXT,
            font=("Segoe UI", 9),
            wraplength=420,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))

        tech_card = tk.Frame(
            signal_row,
            bg="#F2EEFA",
            padx=11,
            pady=8,
            highlightthickness=1,
            highlightbackground="#DCCFF1",
        )
        tech_card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(4, 0),
        )
        tk.Label(
            tech_card,
            text="TECHNOLOGY SIGNALS",
            bg="#F2EEFA",
            fg="#6545A4",
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            tech_card,
            textvariable=self.company_intelligence_vars["technology"],
            bg="#F2EEFA",
            fg=TEXT,
            font=("Segoe UI", 9),
            wraplength=420,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))

        darwill_card = tk.Frame(
            company_tab,
            bg=NAVY,
            padx=12,
            pady=10,
        )
        darwill_card.pack(fill="x", pady=(0, 8))
        tk.Label(
            darwill_card,
            text="RECOMMENDED DARWILL ANGLE",
            bg=NAVY,
            fg="#75B6F5",
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        darwill_angle_value = tk.Label(
            darwill_card,
            textvariable=self.company_intelligence_vars["darwill_angle"],
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI", 9),
            justify="left",
            anchor="nw",
            wraplength=520,
        )
        darwill_angle_value.pack(
            fill="x",
            anchor="w",
            pady=(4, 2),
        )

        def resize_darwill_angle(event):
            darwill_angle_value.configure(
                wraplength=max(220, event.width - 24)
            )

        darwill_card.bind(
            "<Configure>",
            resize_darwill_angle,
            add="+",
        )

        evidence_frame = ttk.Frame(
            company_tab,
            style="Card.TFrame",
        )
        evidence_frame.pack(fill="both", expand=True)
        ttk.Label(
            evidence_frame,
            text="Supporting Evidence and Research",
            style="Card.TLabel",
        ).pack(anchor="w", pady=(0, 4))
        evidence_text_frame = ttk.Frame(
            evidence_frame,
            style="Card.TFrame",
        )
        evidence_text_frame.pack(fill="both", expand=True)
        self.company_intelligence_text = tk.Text(
            evidence_text_frame,
            wrap="word",
            bg=SURFACE_ALT,
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 10),
            padx=12,
            pady=10,
            height=20,
        )
        configure_scrollable_text(
            evidence_text_frame,
            self.company_intelligence_text,
            horizontal=False,
            vertical=True,
            enable_mousewheel=True,
        )
        intelligence_header = ttk.Frame(
            intelligence_tab,
            style="Card.TFrame",
        )
        intelligence_header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            intelligence_header,
            text="Contact Intelligence",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        contact_header_description = ttk.Label(
            intelligence_header,
            text=(
                "Ranking rationale, contactability, verification, and "
                "recommended outreach order."
            ),
            style="Muted.TLabel",
            wraplength=600,
            justify="left",
        )
        contact_header_description.pack(fill="x", anchor="w", pady=(2, 0))
        intelligence_header.bind(
            "<Configure>",
            lambda event: contact_header_description.configure(
                wraplength=max(220, event.width - 12)
            ),
            add="+",
        )

        intelligence_score_row = tk.Frame(
            intelligence_tab,
            bg=SURFACE,
        )
        intelligence_score_row.pack(fill="x", pady=(0, 8))

        self.contact_intelligence_metric_vars = {
            "rank": tk.StringVar(value="—"),
            "decision": tk.StringVar(value="—"),
            "email": tk.StringVar(value="—"),
            "phone": tk.StringVar(value="—"),
        }
        intelligence_specs = [
            ("Outreach Rank", "rank"),
            ("Decision Confidence", "decision"),
            ("Email Confidence", "email"),
            ("Phone Confidence", "phone"),
        ]
        for column, (label, key) in enumerate(intelligence_specs):
            card = tk.Frame(
                intelligence_score_row,
                bg="#F4F8FC",
                padx=10,
                pady=8,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 5, 0),
            )
            tk.Label(
                card,
                text=label.upper(),
                bg="#F4F8FC",
                fg=MUTED,
                font=("Segoe UI Semibold", 7),
            ).pack(anchor="w")
            tk.Label(
                card,
                textvariable=self.contact_intelligence_metric_vars[key],
                bg="#F4F8FC",
                fg=NAVY if key == "rank" else ACCENT,
                font=("Segoe UI Semibold", 15),
            ).pack(anchor="w", pady=(2, 0))
            intelligence_score_row.columnconfigure(column, weight=1)

        quality_row = tk.Frame(
            intelligence_tab,
            bg=SURFACE,
        )
        quality_row.pack(fill="x", pady=(0, 8))
        email_quality = tk.Frame(
            quality_row,
            bg="#ECF8F2",
            padx=11,
            pady=8,
            highlightthickness=1,
            highlightbackground="#CBE8D9",
        )
        email_quality.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 4),
        )
        tk.Label(
            email_quality,
            text="EMAIL INTELLIGENCE",
            bg="#ECF8F2",
            fg=SUCCESS,
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            email_quality,
            textvariable=self.review_email_detail,
            bg="#ECF8F2",
            fg=TEXT,
            justify="left",
            wraplength=430,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(3, 0))

        phone_quality = tk.Frame(
            quality_row,
            bg="#F2EEFA",
            padx=11,
            pady=8,
            highlightthickness=1,
            highlightbackground="#DCCFF1",
        )
        phone_quality.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(4, 0),
        )
        tk.Label(
            phone_quality,
            text="PHONE INTELLIGENCE",
            bg="#F2EEFA",
            fg="#6545A4",
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            phone_quality,
            textvariable=self.review_phone_detail,
            bg="#F2EEFA",
            fg=TEXT,
            justify="left",
            wraplength=430,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(3, 0))

        contact_report_frame = ttk.Frame(
            intelligence_tab,
            style="Card.TFrame",
        )
        contact_report_frame.pack(fill="x", expand=False)
        ttk.Label(
            contact_report_frame,
            text="Ranking Rationale and Contact Brief",
            style="Card.TLabel",
        ).pack(anchor="w", pady=(0, 4))
        contact_text_frame = ttk.Frame(
            contact_report_frame,
            style="Card.TFrame",
        )
        contact_text_frame.pack(fill="x")
        self.contact_intelligence_text = tk.Text(
            contact_text_frame,
            wrap="word",
            bg=SURFACE_ALT,
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 10),
            padx=12,
            pady=10,
            height=20,
        )
        configure_scrollable_text(
            contact_text_frame,
            self.contact_intelligence_text,
            horizontal=False,
            vertical=True,
            enable_mousewheel=True,
        )

        required_company_intelligence_keys = {
            "fit": "Not evaluated",
            "confidence": "—",
            "email_status": "Not evaluated",
            "next_action": "Review company",
            "executive_summary": (
                "Select a company to generate an executive summary."
            ),
        }
        for variable_key, default_value in (
            required_company_intelligence_keys.items()
        ):
            if variable_key not in self.company_intelligence_vars:
                self.company_intelligence_vars[variable_key] = tk.StringVar(
                    value=default_value
                )

        company_intelligence_header = ttk.Frame(
            acquisition_tab,
            style="Card.TFrame",
        )
        company_intelligence_header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            company_intelligence_header,
            text="Company Intelligence Decision",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        company_decision_description = ttk.Label(
            company_intelligence_header,
            text=(
                "Explainable company fit, risk, email state, and recommended "
                "next action from stored research."
            ),
            style="Muted.TLabel",
            wraplength=600,
            justify="left",
        )
        company_decision_description.pack(fill="x", anchor="w", pady=(2, 0))
        company_intelligence_header.bind(
            "<Configure>",
            lambda event: company_decision_description.configure(
                wraplength=max(220, event.width - 12)
            ),
            add="+",
        )

        company_decision_row = tk.Frame(acquisition_tab, bg=SURFACE)
        company_decision_row.pack(fill="x", pady=(0, 8))
        for column, (label, key, background, foreground) in enumerate([
            ("Company Fit", "fit", "#EDF5FC", ACCENT),
            ("Confidence", "confidence", "#FFF4DB", WARNING),
            ("Email State", "email_status", "#F2EEFA", "#6545A4"),
            ("Next Action", "next_action", "#ECF8F2", SUCCESS),
        ]):
            card = tk.Frame(
                company_decision_row,
                bg=background,
                padx=9,
                pady=7,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 4, 0),
            )
            tk.Label(
                card,
                text=label.upper(),
                bg=background,
                fg=MUTED,
                font=("Segoe UI Semibold", 6),
            ).pack(anchor="w")
            tk.Label(
                card,
                textvariable=self.company_intelligence_vars[key],
                bg=background,
                fg=foreground,
                font=("Segoe UI Semibold", 10),
                wraplength=190,
                justify="left",
            ).pack(anchor="w", pady=(2, 0))
            company_decision_row.columnconfigure(column, weight=1)

        email_decision_header = ttk.Frame(
            acquisition_tab,
            style="Card.TFrame",
        )
        email_decision_header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            email_decision_header,
            text="Email Acquisition Decision",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        email_decision_description = ttk.Label(
            email_decision_header,
            text=(
                "Use free/public research first. Recommend a ZoomInfo credit "
                "only when evidence indicates a verified email is available."
            ),
            style="Muted.TLabel",
            wraplength=600,
            justify="left",
        )
        email_decision_description.pack(fill="x", anchor="w", pady=(2, 0))
        email_decision_header.bind(
            "<Configure>",
            lambda event: email_decision_description.configure(
                wraplength=max(220, event.width - 12)
            ),
            add="+",
        )

        email_decision_row = tk.Frame(acquisition_tab, bg=SURFACE)
        email_decision_row.pack(fill="x", pady=(0, 8))
        for column, (label, key, background, foreground) in enumerate([
            ("Public Email", "public_status", "#ECF8F2", SUCCESS),
            ("Pattern", "pattern", "#EDF5FC", ACCENT),
            ("Confidence", "confidence", "#FFF4DB", WARNING),
            ("ZoomInfo", "zoominfo", "#F2EEFA", "#6545A4"),
            ("Recommendation", "recommendation", "#FDEEEE", ERROR),
        ]):
            card = tk.Frame(
                email_decision_row,
                bg=background,
                padx=9,
                pady=7,
                highlightthickness=1,
                highlightbackground=BORDER,
            )
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=(0 if column == 0 else 4, 0),
            )
            tk.Label(
                card,
                text=label.upper(),
                bg=background,
                fg=MUTED,
                font=("Segoe UI Semibold", 6),
            ).pack(anchor="w")
            tk.Label(
                card,
                textvariable=self.email_intelligence_vars[key],
                bg=background,
                fg=foreground,
                font=("Segoe UI Semibold", 10),
                wraplength=145,
                justify="left",
            ).pack(anchor="w", pady=(2, 0))
            email_decision_row.columnconfigure(column, weight=1)

        acquisition_header = ttk.Frame(
            acquisition_tab,
            style="Card.TFrame",
        )
        acquisition_header.pack(fill="x", pady=(0, 8))
        acquisition_header.columnconfigure(0, weight=1)

        acquisition_title_area = ttk.Frame(
            acquisition_header,
            style="Card.TFrame",
        )
        acquisition_title_area.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8),
        )
        ttk.Label(
            acquisition_title_area,
            text="Contact Acquisition Report",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        acquisition_header_description = ttk.Label(
            acquisition_title_area,
            text=(
                "Source trail for the person, email, phone, "
                "and any predicted address."
            ),
            style="Muted.TLabel",
            wraplength=420,
            justify="left",
        )
        acquisition_header_description.pack(
            fill="x",
            anchor="w",
            pady=(2, 0),
        )

        acquisition_actions = ttk.Frame(
            acquisition_header,
            style="Card.TFrame",
        )
        acquisition_actions.grid(row=0, column=1, sticky="ne")
        ttk.Button(
            acquisition_actions,
            text="Retry ZoomInfo Email",
            command=self._retry_selected_contact_enrichment,
            style="Primary.TButton",
        ).pack(side="left")
        ttk.Button(
            acquisition_actions,
            text="MCP Explorer",
            command=self._open_mcp_explorer,
            style="Secondary.TButton",
        ).pack(side="left", padx=(8, 0))

        acquisition_title_area.bind(
            "<Configure>",
            lambda event: acquisition_header_description.configure(
                wraplength=max(180, event.width - 8)
            ),
            add="+",
        )

        acquisition_report_frame = ttk.Frame(
            acquisition_tab,
            style="Card.TFrame",
        )
        acquisition_report_frame.pack(fill="x")
        self.contact_acquisition_text = tk.Text(
            acquisition_report_frame,
            wrap="word",
            bg=SURFACE_ALT,
            fg=TEXT,
            relief="solid",
            bd=1,
            font=("Segoe UI", 10),
            padx=12,
            pady=10,
            height=28,
        )
        configure_scrollable_text(
            acquisition_report_frame,
            self.contact_acquisition_text,
            horizontal=False,
            vertical=True,
            enable_mousewheel=True,
        )

        review_actions = ttk.Frame(
            review_tabs_pane, style="Card.TFrame"
        )
        review_actions.pack(fill="x", pady=(6, 0))
        ttk.Button(
            review_actions,
            text="Save Edits",
            style="Secondary.TButton",
            command=self._save_queue_edits,
        ).pack(side="left")
        ttk.Button(
            review_actions,
            text="View Contact Intelligence",
            style="Secondary.TButton",
            command=self._show_contact_intelligence_tab,
        ).pack(side="right")
        ttk.Button(
            review_actions,
            text="View Acquisition Report",
            style="Secondary.TButton",
            command=self._show_contact_acquisition_tab,
        ).pack(side="right", padx=(0, 8))

        integration = ttk.LabelFrame(
            deal_desk_tab, text="HubSpot and Sequence Control", padding=10
        )
        integration.pack(fill="x", pady=(12, 0))

        self._row(
            integration, 0, "HubSpot access token / service key",
            self.hubspot_token, show="•",
        )
        self._row(
            integration, 1, "Sequence sender email",
            self.hubspot_sender_email,
        )

        ttk.Label(
            integration,
            text="Sequence",
            style="Card.TLabel",
        ).grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        self.sequence_combo = ttk.Combobox(
            integration,
            textvariable=self.selected_sequence,
            state="readonly",
        )
        self.sequence_combo.grid(row=2, column=1, sticky="ew", pady=5)

        hubspot_buttons = ttk.Frame(
            integration, style="Card.TFrame"
        )
        hubspot_buttons.grid(
            row=3, column=1, sticky="w", pady=(10, 0)
        )
        ttk.Button(
            hubspot_buttons,
            text="Test HubSpot",
            style="Secondary.TButton",
            command=self._test_hubspot,
        ).pack(side="left")
        ttk.Button(
            hubspot_buttons,
            text="Load Sequences",
            style="Secondary.TButton",
            command=self._load_hubspot_sequences,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            hubspot_buttons,
            text="Sync Approved Records",
            style="Primary.TButton",
            command=self._sync_approved_to_hubspot,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            hubspot_buttons,
            text="Final Review & Enroll",
            style="Primary.TButton",
            command=self._confirm_and_enroll,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            hubspot_buttons,
            text="Export Approved CSV",
            style="Secondary.TButton",
            command=self._export_approved_hubspot_csv,
        ).pack(side="left", padx=(8, 0))

        ttk.Label(
            integration,
            text=(
                "Syncing creates/associates records and adds a review note. "
                "It does not send email. Sequence enrollment is a separate "
                "final-confirmation action because enrollment can begin outbound activity."
            ),
            style="Muted.TLabel",
            wraplength=1050,
        ).grid(
            row=4, column=0, columnspan=3,
            sticky="w", pady=(10, 0)
        )
        integration.columnconfigure(1, weight=1)

        # Deliverability Center
        delivery_header = tk.Frame(
            deliverability_tab, bg=NAVY, padx=20, pady=15
        )
        delivery_header.pack(fill="x", pady=(0, 12))
        tk.Label(
            delivery_header, text="OUTREACH QUALITY",
            bg=NAVY, fg="#75B6F5",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            delivery_header, text="Deliverability Center",
            bg=NAVY, fg=WHITE,
            font=("Segoe UI Semibold", 19),
        ).pack(anchor="w", pady=(2, 1))
        tk.Label(
            delivery_header,
            text="Reduce avoidable inbox risk, improve messaging quality, and record real-world outcomes.",
            bg=NAVY, fg="#CFE3F6",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        delivery_pane = ttk.Panedwindow(
            deliverability_tab, orient="horizontal"
        )
        delivery_pane.pack(fill="both", expand=True)

        score_frame = ttk.LabelFrame(
            delivery_pane, text="Inbox Readiness", padding=12
        )
        detail_frame = ttk.LabelFrame(
            delivery_pane, text="Analysis and Recommended Changes", padding=12
        )
        delivery_pane.add(score_frame, weight=2)
        delivery_pane.add(detail_frame, weight=4)

        self.delivery_company = tk.StringVar(value="Select an item in Deal Desk")
        self.delivery_contact = tk.StringVar(value="")
        self.delivery_overall = tk.StringVar(value="—")
        self.delivery_risk = tk.StringVar(value="Not analyzed")
        self.delivery_subject_score = tk.StringVar(value="—")
        self.delivery_personalization_score = tk.StringVar(value="—")
        self.delivery_content_score = tk.StringVar(value="—")
        self.delivery_tone_score = tk.StringVar(value="—")
        self.delivery_format_score = tk.StringVar(value="—")
        self.delivery_auth_score = tk.StringVar(value="—")
        self.delivery_word_count = tk.StringVar(value="—")
        self.delivery_link_count = tk.StringVar(value="—")
        self.delivery_cta_count = tk.StringVar(value="—")
        self.delivery_alt_subject_1 = tk.StringVar()
        self.delivery_alt_subject_2 = tk.StringVar()
        self.delivery_alt_subject_3 = tk.StringVar()

        ttk.Label(
            score_frame,
            textvariable=self.delivery_company,
            style="SectionTitle.TLabel",
            wraplength=340,
        ).pack(anchor="w")
        ttk.Label(
            score_frame,
            textvariable=self.delivery_contact,
            style="Muted.TLabel",
            wraplength=340,
        ).pack(anchor="w", pady=(2, 12))

        score_card = tk.Frame(
            score_frame,
            bg=NAVY,
            padx=18,
            pady=16,
        )
        score_card.pack(fill="x", pady=(0, 12))
        tk.Label(
            score_card,
            textvariable=self.delivery_overall,
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 30),
        ).pack()
        tk.Label(
            score_card,
            text="INBOX READINESS",
            bg=NAVY,
            fg="#B9D8EF",
            font=("Segoe UI Semibold", 8),
        ).pack()
        tk.Label(
            score_card,
            textvariable=self.delivery_risk,
            bg=NAVY,
            fg="#8ED4AD",
            font=("Segoe UI Semibold", 10),
        ).pack(pady=(5, 0))

        metrics = [
            ("Subject quality", self.delivery_subject_score),
            ("Personalization", self.delivery_personalization_score),
            ("Content risk", self.delivery_content_score),
            ("Human tone", self.delivery_tone_score),
            ("Formatting", self.delivery_format_score),
            ("Sender authentication", self.delivery_auth_score),
            ("Words", self.delivery_word_count),
            ("Links", self.delivery_link_count),
            ("CTA signals", self.delivery_cta_count),
        ]
        for label, variable in metrics:
            row = tk.Frame(score_frame, bg=SURFACE)
            row.pack(fill="x", pady=3)
            tk.Label(
                row, text=label, bg=SURFACE, fg=MUTED,
                font=("Segoe UI", 9),
            ).pack(side="left")
            tk.Label(
                row, textvariable=variable, bg=SURFACE, fg=NAVY,
                font=("Segoe UI Semibold", 9),
            ).pack(side="right")

        sender_frame = ttk.LabelFrame(
            score_frame, text="Sender Health Checklist", padding=8
        )
        sender_frame.pack(fill="x", pady=(14, 0))
        ttk.Checkbutton(
            sender_frame, text="SPF confirmed",
            variable=self.spf_confirmed,
            command=self._save_deliverability_settings,
        ).pack(anchor="w")
        ttk.Checkbutton(
            sender_frame, text="DKIM confirmed",
            variable=self.dkim_confirmed,
            command=self._save_deliverability_settings,
        ).pack(anchor="w")
        ttk.Checkbutton(
            sender_frame, text="DMARC confirmed",
            variable=self.dmarc_confirmed,
            command=self._save_deliverability_settings,
        ).pack(anchor="w")

        threshold_row = ttk.Frame(sender_frame, style="Card.TFrame")
        threshold_row.pack(fill="x", pady=(8, 0))
        ttk.Label(
            threshold_row,
            text="Minimum enrollment score",
            style="Card.TLabel",
        ).pack(side="left")
        ttk.Spinbox(
            threshold_row,
            from_=60, to=100,
            textvariable=self.minimum_inbox_score,
            width=6,
            command=self._save_deliverability_settings,
        ).pack(side="right")

        ttk.Checkbutton(
            sender_frame,
            text="Allow explicit low-score override",
            variable=self.allow_low_score_override,
        ).pack(anchor="w", pady=(8, 0))

        ttk.Label(
            detail_frame, text="Issues",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        self.delivery_issues = tk.Text(
            detail_frame, height=7, wrap="word",
            bg="#FFF9F7", fg=ERROR, relief="flat", bd=0,
            highlightthickness=1, highlightbackground="#F0C9C0",
            font=("Segoe UI", 9), padx=10, pady=8,
        )
        self.delivery_issues.pack(fill="x", pady=(5, 12))

        ttk.Label(
            detail_frame, text="Recommended changes",
            style="SectionTitle.TLabel",
        ).pack(anchor="w")
        self.delivery_recommendations = tk.Text(
            detail_frame, height=7, wrap="word",
            bg="#F5FAFF", fg=NAVY, relief="flat", bd=0,
            highlightthickness=1, highlightbackground="#C8DDF0",
            font=("Segoe UI", 9), padx=10, pady=8,
        )
        self.delivery_recommendations.pack(fill="x", pady=(5, 12))

        subjects = ttk.LabelFrame(
            detail_frame, text="Safer Subject Alternatives", padding=8
        )
        subjects.pack(fill="x", pady=(0, 12))
        for index, variable in enumerate([
            self.delivery_alt_subject_1,
            self.delivery_alt_subject_2,
            self.delivery_alt_subject_3,
        ], start=1):
            row = ttk.Frame(subjects, style="Card.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Label(
                row, textvariable=variable, style="Card.TLabel"
            ).pack(side="left", fill="x", expand=True)
            ttk.Button(
                row,
                text=f"Use {index}",
                style="Secondary.TButton",
                command=lambda idx=index: self._apply_subject_alternative(idx),
            ).pack(side="right")

        delivery_actions = ttk.Frame(
            detail_frame, style="Card.TFrame"
        )
        delivery_actions.pack(fill="x", pady=(0, 12))
        ttk.Button(
            delivery_actions,
            text="Analyze Selected Draft",
            style="Primary.TButton",
            command=self._analyze_selected_delivery,
        ).pack(side="left")
        ttk.Button(
            delivery_actions,
            text="Apply Safer Rewrite",
            style="Secondary.TButton",
            command=self._apply_safer_rewrite,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            delivery_actions,
            text="Save & Return to Deal Desk",
            style="Secondary.TButton",
            command=self._save_delivery_to_queue,
        ).pack(side="left", padx=(8, 0))

        outcome_frame = ttk.LabelFrame(
            detail_frame, text="Performance Learning", padding=8
        )
        outcome_frame.pack(fill="x")
        ttk.Label(
            outcome_frame,
            text=(
                "After sending, record the result so future versions can compare "
                "subjects, strategies, and outcomes."
            ),
            style="Muted.TLabel",
            wraplength=650,
        ).pack(anchor="w", pady=(0, 6))
        outcome_row = ttk.Frame(outcome_frame, style="Card.TFrame")
        outcome_row.pack(fill="x")
        ttk.Combobox(
            outcome_row,
            textvariable=self.delivery_outcome,
            state="readonly",
            values=[
                "Not Sent", "Sent", "Opened", "Replied",
                "Meeting Booked", "Bounced", "Spam/Junk",
                "No Response",
            ],
            width=22,
        ).pack(side="left")
        ttk.Button(
            outcome_row,
            text="Save Outcome",
            style="Secondary.TButton",
            command=self._save_delivery_outcome,
        ).pack(side="left", padx=(8, 0))

        master_header = ttk.Frame(master_tab, style="Card.TFrame")
        master_header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            master_header,
            text="Master Prospect Database",
            style="SectionTitle.TLabel",
        ).pack(side="left")
        ttk.Label(
            master_header,
            text=(
                "One persistent lifecycle record per company across every search run."
            ),
            style="Muted.TLabel",
        ).pack(side="left", padx=(12, 0))

        permanent_master = ttk.LabelFrame(
            master_tab,
            text="Permanent Master CSV",
            padding=10,
        )
        permanent_master.pack(fill="x", pady=(0, 10))

        master_csv_status = tk.Frame(
            permanent_master,
            bg=SURFACE_ALT,
            highlightthickness=1,
            highlightbackground=BORDER,
            padx=12,
            pady=9,
        )
        master_csv_status.grid(
            row=0,
            column=0,
            columnspan=4,
            sticky="ew",
            pady=(0, 8),
        )
        tk.Label(
            master_csv_status,
            textvariable=self.master_csv_loaded_file,
            bg=SURFACE_ALT,
            fg=NAVY,
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            master_csv_status,
            text=(
                "This file persists across application versions and is "
                "backed up before each write."
            ),
            bg=SURFACE_ALT,
            fg=MUTED,
            font=("Segoe UI", 8),
            anchor="w",
        ).pack(fill="x", pady=(2, 0))

        ttk.Label(
            permanent_master,
            text="File",
            style="Card.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(
            permanent_master,
            textvariable=self.master_csv_path,
        ).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(
            permanent_master,
            text="Browse",
            style="Secondary.TButton",
            command=self._browse_master_csv,
        ).grid(row=1, column=2, padx=(8, 0))
        ttk.Button(
            permanent_master,
            text="Load / Import",
            style="Primary.TButton",
            command=self._load_master_csv,
        ).grid(row=1, column=3, padx=(8, 0))

        ttk.Checkbutton(
            permanent_master,
            text="Automatically update this CSV after every completed run",
            variable=self.master_csv_auto_sync,
        ).grid(
            row=2,
            column=1,
            sticky="w",
            pady=(5, 0),
        )
        ttk.Button(
            permanent_master,
            text="Sync Now",
            style="Secondary.TButton",
            command=self._sync_master_csv_now,
        ).grid(row=2, column=2, padx=(8, 0), pady=(5, 0))
        ttk.Button(
            permanent_master,
            text="Open Folder",
            style="Secondary.TButton",
            command=self._open_master_csv_folder,
        ).grid(row=2, column=3, padx=(8, 0), pady=(5, 0))

        ttk.Label(
            permanent_master,
            text="Records",
            style="Card.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            permanent_master,
            textvariable=self.master_csv_company_count,
            style="Card.TLabel",
        ).grid(row=3, column=1, sticky="w", pady=(8, 0))
        ttk.Label(
            permanent_master,
            text="Last updated",
            style="Card.TLabel",
        ).grid(row=3, column=2, sticky="e", pady=(8, 0))
        ttk.Label(
            permanent_master,
            textvariable=self.master_csv_last_updated,
            style="Card.TLabel",
        ).grid(row=3, column=3, sticky="w", padx=(8, 0), pady=(8, 0))
        permanent_master.columnconfigure(1, weight=1)

        master_filters = ttk.Frame(master_tab, style="Card.TFrame")
        master_filters.pack(fill="x", pady=(0, 10))
        ttk.Label(
            master_filters, text="Status", style="Card.TLabel"
        ).pack(side="left")
        status_combo = ttk.Combobox(
            master_filters,
            textvariable=self.master_status_filter,
            state="readonly",
            values=["All"] + PROSPECT_STATUSES,
            width=22,
        )
        status_combo.pack(side="left", padx=(6, 14))
        status_combo.bind(
            "<<ComboboxSelected>>",
            lambda _event: self._refresh_master_database(),
        )
        ttk.Label(
            master_filters, text="Search", style="Card.TLabel"
        ).pack(side="left")
        master_search_entry = ttk.Entry(
            master_filters,
            textvariable=self.master_search,
            width=34,
        )
        master_search_entry.pack(side="left", padx=(6, 8))
        master_search_entry.bind(
            "<Return>",
            lambda _event: self._refresh_master_database(),
        )
        ttk.Button(
            master_filters,
            text="Refresh",
            style="Secondary.TButton",
            command=self._refresh_master_database,
        ).pack(side="left")
        ttk.Button(
            master_filters,
            text="Export Master CSV",
            style="Secondary.TButton",
            command=self._export_master_database,
        ).pack(side="left", padx=(8, 0))

        self.master_count_vars = {
            "Total": tk.StringVar(value="0"),
            "Qualified": tk.StringVar(value="0"),
            "Rejected": tk.StringVar(value="0"),
            "Approved": tk.StringVar(value="0"),
            "In Sequence": tk.StringVar(value="0"),
            "Meeting": tk.StringVar(value="0"),
            "Customer": tk.StringVar(value="0"),
        }
        master_cards = ttk.Frame(master_tab, style="Card.TFrame")
        master_cards.pack(fill="x", pady=(0, 10))
        for col, (label, variable) in enumerate(self.master_count_vars.items()):
            card = tk.Frame(
                master_cards,
                bg=SURFACE,
                highlightthickness=1,
                highlightbackground=BORDER,
                padx=12,
                pady=8,
            )
            card.grid(row=0, column=col, sticky="nsew", padx=3)
            tk.Label(
                card,
                textvariable=variable,
                bg=SURFACE,
                fg=NAVY,
                font=("Segoe UI Semibold", 16),
            ).pack()
            tk.Label(
                card,
                text=label,
                bg=SURFACE,
                fg=MUTED,
                font=("Segoe UI Semibold", 8),
            ).pack()
            master_cards.columnconfigure(col, weight=1)

        master_pane = ttk.Panedwindow(master_tab, orient="vertical")
        master_pane.pack(fill="both", expand=True)

        master_table_frame = ttk.Frame(master_pane, style="Card.TFrame")
        master_detail_frame = ttk.LabelFrame(
            master_pane, text="Lifecycle Update", padding=10
        )
        master_pane.add(master_table_frame, weight=5)
        master_pane.add(master_detail_frame, weight=2)

        master_columns = (
            "company", "domain", "state", "trade", "revenue",
            "employees", "status", "decision", "fit", "reviewed",
            "next_review", "hubspot", "sequence",
        )
        self.master_tree = ttk.Treeview(
            master_table_frame,
            columns=master_columns,
            show="headings",
        )
        master_headings = {
            "company": "Company",
            "domain": "Domain",
            "state": "State",
            "trade": "Trade",
            "revenue": "Revenue",
            "employees": "Employees",
            "status": "Lifecycle",
            "decision": "Decision",
            "fit": "Fit",
            "reviewed": "Last Reviewed",
            "next_review": "Next Review",
            "hubspot": "HubSpot ID",
            "sequence": "Sequence",
        }
        master_widths = {
            "company": 180, "domain": 150, "state": 55, "trade": 110,
            "revenue": 90, "employees": 75, "status": 115,
            "decision": 80, "fit": 55, "reviewed": 115,
            "next_review": 115, "hubspot": 95, "sequence": 95,
        }
        for name in master_columns:
            self.master_tree.heading(name, text=master_headings[name])
            self.master_tree.column(
                name, width=master_widths[name], stretch=True
            )
        master_scroll_y = ttk.Scrollbar(
            master_table_frame, command=self.master_tree.yview
        )
        master_scroll_x = ttk.Scrollbar(
            master_table_frame,
            orient="horizontal",
            command=self.master_tree.xview,
        )
        self.master_tree.configure(
            yscrollcommand=master_scroll_y.set,
            xscrollcommand=master_scroll_x.set,
        )
        self.master_tree.grid(row=0, column=0, sticky="nsew")
        master_scroll_y.grid(row=0, column=1, sticky="ns")
        master_scroll_x.grid(row=1, column=0, sticky="ew")
        master_table_frame.rowconfigure(0, weight=1)
        master_table_frame.columnconfigure(0, weight=1)
        self.master_tree.bind(
            "<<TreeviewSelect>>",
            lambda _event: self._load_master_detail(),
        )

        self.master_selected_key = tk.StringVar()
        self.master_selected_company = tk.StringVar()
        self.master_selected_status = tk.StringVar()
        self.master_selected_notes = tk.StringVar()

        ttk.Label(
            master_detail_frame,
            textvariable=self.master_selected_company,
            style="SectionTitle.TLabel",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Label(
            master_detail_frame,
            text="Lifecycle status",
            style="Card.TLabel",
        ).grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Combobox(
            master_detail_frame,
            textvariable=self.master_selected_status,
            state="readonly",
            values=PROSPECT_STATUSES,
            width=24,
        ).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(
            master_detail_frame,
            text="Notes",
            style="Card.TLabel",
        ).grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(
            master_detail_frame,
            textvariable=self.master_selected_notes,
        ).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(
            master_detail_frame,
            text="Update Lifecycle",
            style="Primary.TButton",
            command=self._update_master_status,
        ).grid(row=1, column=2, rowspan=2, padx=(12, 0))
        master_detail_frame.columnconfigure(1, weight=1)

        hubspot_header = tk.Frame(
            hubspot_workspace_tab, bg=NAVY, padx=20, pady=14
        )
        hubspot_header.pack(fill="x", pady=(0, 12))
        tk.Label(
            hubspot_header, text="CRM CONTROL",
            bg=NAVY, fg="#75B6F5",
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w")
        tk.Label(
            hubspot_header, text="HubSpot Workspace",
            bg=NAVY, fg=WHITE,
            font=("Segoe UI Semibold", 19),
        ).pack(anchor="w", pady=(2, 1))
        tk.Label(
            hubspot_header,
            text=(
                "Connect securely, synchronize reviewed records, and enroll "
                "only after an explicit final approval."
            ),
            bg=NAVY, fg="#CFE3F6",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        hubspot_kpis = tk.Frame(hubspot_workspace_tab, bg=SURFACE)
        hubspot_kpis.pack(fill="x", pady=(0, 12))
        for column, (label, key, background, foreground) in enumerate([
            ("Approved", "approved", "#EAF3FB", ACCENT),
            ("Ready to Sync", "ready", "#FFF4DB", WARNING),
            ("Synced", "synced", "#E9F7F0", SUCCESS),
            ("Enrolled", "enrolled", "#F1ECFB", "#6545A4"),
            ("Failed", "failed", "#FDEEEE", ERROR),
        ]):
            card = tk.Frame(
                hubspot_kpis, bg=background, padx=12, pady=8,
                highlightthickness=1, highlightbackground=BORDER,
            )
            card.grid(
                row=0, column=column, sticky="nsew",
                padx=(0 if column == 0 else 5, 0),
            )
            tk.Label(
                card, textvariable=self.hubspot_kpi_vars[key],
                bg=background, fg=foreground,
                font=("Segoe UI Semibold", 17),
            ).pack(anchor="w")
            tk.Label(
                card, text=label, bg=background, fg=TEXT,
                font=("Segoe UI Semibold", 8),
            ).pack(anchor="w", pady=(2, 0))
            hubspot_kpis.columnconfigure(column, weight=1)

        hubspot_body = ttk.Panedwindow(
            hubspot_workspace_tab, orient="horizontal"
        )
        hubspot_body.pack(fill="both", expand=True)

        connection_panel = ttk.LabelFrame(
            hubspot_body, text="Connection and Sequence", padding=12
        )
        records_panel = ttk.LabelFrame(
            hubspot_body, text="Approved CRM Queue", padding=10
        )
        hubspot_body.add(connection_panel, weight=2)
        hubspot_body.add(records_panel, weight=4)

        self._row(
            connection_panel, 0, "Private app access token",
            self.hubspot_token, show="•",
        )
        self._row(
            connection_panel, 1, "Sequence sender email",
            self.hubspot_sender_email,
        )
        ttk.Label(
            connection_panel, text="Sequence", style="Card.TLabel"
        ).grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        self.hubspot_workspace_sequence_combo = ttk.Combobox(
            connection_panel,
            textvariable=self.selected_sequence,
            state="readonly",
        )
        self.hubspot_workspace_sequence_combo.grid(
            row=2, column=1, sticky="ew", pady=5
        )
        connection_panel.columnconfigure(1, weight=1)

        status_card = tk.Frame(
            connection_panel, bg="#F4F8FC", padx=10, pady=9,
            highlightthickness=1, highlightbackground=BORDER,
        )
        status_card.grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(10, 8)
        )
        tk.Label(
            status_card, text="CONNECTION STATUS",
            bg="#F4F8FC", fg=MUTED,
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            status_card, textvariable=self.hubspot_connection_status,
            bg="#F4F8FC", fg=NAVY,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", pady=(3, 0))

        connection_buttons = ttk.Frame(connection_panel, style="Card.TFrame")
        connection_buttons.grid(
            row=4, column=0, columnspan=2, sticky="ew"
        )
        ttk.Button(
            connection_buttons, text="Save Securely",
            style="Secondary.TButton", command=self._save_credentials,
        ).pack(side="left")
        ttk.Button(
            connection_buttons, text="Test Connection",
            style="Secondary.TButton", command=self._test_hubspot,
        ).pack(side="left", padx=(7, 0))
        ttk.Button(
            connection_buttons, text="Load Sequences",
            style="Secondary.TButton",
            command=self._load_hubspot_sequences,
        ).pack(side="left", padx=(7, 0))

        safety = tk.Frame(
            connection_panel, bg="#FFF8E8", padx=10, pady=9,
            highlightthickness=1, highlightbackground="#EBD59B",
        )
        safety.grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(12, 0)
        )
        tk.Label(
            safety, text="CONTROL POLICY", bg="#FFF8E8", fg=WARNING,
            font=("Segoe UI Semibold", 7),
        ).pack(anchor="w")
        tk.Label(
            safety,
            text=(
                "Sync creates or reuses company/contact records and adds a "
                "research note. It does not send email. Enrollment remains "
                "a separate final-confirmation action."
            ),
            bg="#FFF8E8", fg=TEXT, wraplength=390, justify="left",
            font=("Segoe UI", 8),
        ).pack(anchor="w", pady=(3, 0))

        records_toolbar = ttk.Frame(records_panel, style="Card.TFrame")
        records_toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(
            records_toolbar, text="Refresh Queue",
            style="Secondary.TButton",
            command=self._refresh_hubspot_workspace,
        ).pack(side="left")
        ttk.Button(
            records_toolbar, text="Sync Selected",
            style="Primary.TButton",
            command=self._sync_selected_to_hubspot,
        ).pack(side="left", padx=(7, 0))
        ttk.Button(
            records_toolbar, text="Sync All Approved",
            style="Primary.TButton",
            command=self._sync_approved_to_hubspot,
        ).pack(side="left", padx=(7, 0))
        ttk.Button(
            records_toolbar, text="Final Review & Enroll",
            style="Primary.TButton",
            command=self._confirm_and_enroll,
        ).pack(side="right")

        hubspot_tree_frame = ttk.Frame(records_panel, style="Card.TFrame")
        hubspot_tree_frame.pack(fill="both", expand=True)
        cols = ("company", "contact", "email", "review", "hubspot", "sequence")
        self.hubspot_workspace_tree = ttk.Treeview(
            hubspot_tree_frame, columns=cols, show="headings"
        )
        headings = {
            "company": "Company", "contact": "Contact", "email": "Email",
            "review": "Review", "hubspot": "HubSpot", "sequence": "Sequence",
        }
        widths = {
            "company": 180, "contact": 190, "email": 210,
            "review": 95, "hubspot": 105, "sequence": 120,
        }
        for name in cols:
            self.hubspot_workspace_tree.heading(name, text=headings[name])
            self.hubspot_workspace_tree.column(
                name, width=widths[name], minwidth=80
            )
        configure_scrollable_tree(
            hubspot_tree_frame, self.hubspot_workspace_tree,
            horizontal=True, vertical=True, enable_mousewheel=True,
        )

        message_bar = tk.Frame(
            hubspot_workspace_tab, bg="#F4F8FC", padx=10, pady=7,
            highlightthickness=1, highlightbackground=BORDER,
        )
        message_bar.pack(fill="x", pady=(10, 0))
        tk.Label(
            message_bar, textvariable=self.hubspot_workspace_message,
            bg="#F4F8FC", fg=TEXT, anchor="w",
            font=("Segoe UI", 8),
        ).pack(fill="x")

        index_header = ttk.Frame(hubspot_index_tab, style="Card.TFrame")
        index_header.pack(fill="x", pady=(0, 12))
        ttk.Label(
            index_header,
            text="HubSpot CSV Index",
            style="SectionTitle.TLabel",
        ).pack(side="left")
        ttk.Label(
            index_header,
            text=(
                "Your exported HubSpot companies and contacts are the primary "
                "source of truth for CRM duplicate prevention."
            ),
            style="Muted.TLabel",
        ).pack(side="left", padx=(12, 0))

        self.hubspot_index_company_count = tk.StringVar(value="0")
        self.hubspot_index_contact_count = tk.StringVar(value="0")
        self.hubspot_index_last_import = tk.StringVar(value="No imports yet")
        self.hubspot_index_duplicates_prevented = tk.StringVar(value="0")

        index_cards = ttk.Frame(hubspot_index_tab, style="Card.TFrame")
        index_cards.pack(fill="x", pady=(0, 12))
        for column, (label, variable) in enumerate([
            ("Companies Indexed", self.hubspot_index_company_count),
            ("Contacts Indexed", self.hubspot_index_contact_count),
            ("Last Import", self.hubspot_index_last_import),
            ("Duplicates Prevented This Run", self.hubspot_index_duplicates_prevented),
        ]):
            card = tk.Frame(
                index_cards,
                bg=SURFACE,
                highlightthickness=1,
                highlightbackground=BORDER,
                padx=14,
                pady=10,
            )
            card.grid(row=0, column=column, sticky="nsew", padx=4)
            tk.Label(
                card, textvariable=variable, bg=SURFACE, fg=NAVY,
                font=("Segoe UI Semibold", 15),
                wraplength=220,
            ).pack()
            tk.Label(
                card, text=label, bg=SURFACE, fg=MUTED,
                font=("Segoe UI Semibold", 8),
            ).pack()
            index_cards.columnconfigure(column, weight=1)

        import_frame = ttk.LabelFrame(
            hubspot_index_tab, text="Import or Refresh HubSpot Exports", padding=12
        )
        import_frame.pack(fill="x", pady=(0, 12))
        ttk.Button(
            import_frame,
            text="Import Companies CSV",
            style="Primary.TButton",
            command=lambda: self._import_hubspot_csv("companies", "merge"),
        ).pack(side="left")
        ttk.Button(
            import_frame,
            text="Import Contacts CSV",
            style="Primary.TButton",
            command=lambda: self._import_hubspot_csv("contacts", "merge"),
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            import_frame,
            text="Replace Companies Index",
            style="Secondary.TButton",
            command=lambda: self._import_hubspot_csv("companies", "replace"),
        ).pack(side="left", padx=(18, 0))
        ttk.Button(
            import_frame,
            text="Replace Contacts Index",
            style="Secondary.TButton",
            command=lambda: self._import_hubspot_csv("contacts", "replace"),
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            import_frame,
            text="Clear Entire Index",
            style="Secondary.TButton",
            command=self._clear_hubspot_index,
        ).pack(side="right")

        ttk.Label(
            hubspot_index_tab,
            text=(
                "Merge adds new records and updates matching records. Replace clears "
                "that index first and rebuilds it from the selected full export. "
                "Use Merge for incremental exports and Replace for a fresh complete export."
            ),
            style="Muted.TLabel",
            wraplength=1050,
        ).pack(anchor="w", pady=(0, 10))

        history_frame = ttk.LabelFrame(
            hubspot_index_tab, text="Import History", padding=8
        )
        history_frame.pack(fill="both", expand=True)
        index_columns = (
            "date", "file", "type", "mode",
            "read", "new", "updated", "skipped",
        )
        self.hubspot_index_tree = ttk.Treeview(
            history_frame, columns=index_columns, show="headings"
        )
        headings = {
            "date": "Imported At", "file": "File", "type": "Type",
            "mode": "Mode", "read": "Rows Read", "new": "New",
            "updated": "Updated", "skipped": "Skipped",
        }
        widths = {
            "date": 145, "file": 280, "type": 90, "mode": 80,
            "read": 80, "new": 70, "updated": 80, "skipped": 75,
        }
        for name in index_columns:
            self.hubspot_index_tree.heading(name, text=headings[name])
            self.hubspot_index_tree.column(
                name, width=widths[name], stretch=True
            )
        index_scroll = ttk.Scrollbar(
            history_frame, command=self.hubspot_index_tree.yview
        )
        self.hubspot_index_tree.configure(
            yscrollcommand=index_scroll.set
        )
        self.hubspot_index_tree.pack(
            side="left", fill="both", expand=True
        )
        index_scroll.pack(side="right", fill="y")

        history_controls = ttk.Frame(history_tab, style="Card.TFrame")
        history_controls.pack(fill="x", pady=(0, 8))
        ttk.Button(
            history_controls, text="Refresh", style="Secondary.TButton",
            command=self._refresh_history,
        ).pack(side="left")
        ttk.Button(
            history_controls, text="Open Selected Workbook",
            style="Blue.TButton", command=self._open_selected_history,
        ).pack(side="left", padx=(8, 0))
        history_frame = ttk.Frame(history_tab, style="Card.TFrame")
        history_frame.pack(fill="both", expand=True)
        columns = (
            "started", "profile", "states", "trades", "qualified",
            "contacts", "reviewed", "status", "output",
        )
        self.history_tree = ttk.Treeview(
            history_frame, columns=columns, show="headings",
        )
        headings = {
            "started": "Started", "profile": "Profile", "states": "States",
            "trades": "Trades", "qualified": "Qualified",
            "contacts": "Contacts", "reviewed": "Reviewed",
            "status": "Status", "output": "Workbook",
        }
        for name in columns:
            self.history_tree.heading(name, text=headings[name])
            self.history_tree.column(name, width=120, stretch=True)
        self.history_tree.column("output", width=280)
        history_scroll = ttk.Scrollbar(
            history_frame, command=self.history_tree.yview
        )
        self.history_tree.configure(yscrollcommand=history_scroll.set)
        self.history_tree.pack(side="left", fill="both", expand=True)
        history_scroll.pack(side="right", fill="y")

        learning_header = ttk.Frame(learning_tab, style="Card.TFrame")
        learning_header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            learning_header,
            text="Teach the application why companies should be rejected",
            style="Card.TLabel",
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left")
        ttk.Button(
            learning_header,
            text="Refresh",
            style="Secondary.TButton",
            command=self._refresh_learning,
        ).pack(side="right")

        learning_form = ttk.LabelFrame(
            learning_tab, text="Add rejection feedback", padding=12
        )
        learning_form.pack(fill="x", pady=(0, 12))
        self.feedback_company = tk.StringVar()
        self.feedback_website = tk.StringVar()
        self.feedback_reason = tk.StringVar(value="Commercial only")
        self.feedback_notes = tk.StringVar()

        self._row(learning_form, 0, "Company name", self.feedback_company)
        self._row(learning_form, 1, "Website", self.feedback_website)
        ttk.Label(
            learning_form, text="Rejection reason", style="Card.TLabel"
        ).grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        ttk.Combobox(
            learning_form,
            textvariable=self.feedback_reason,
            state="readonly",
            values=[
                "Commercial only", "Franchise", "Too small", "Multi-state",
                "Manufacturer", "Distributor", "Retailer", "Wrong trade",
                "Private-equity roll-up", "Already contacted", "Other",
            ],
        ).grid(row=2, column=1, sticky="ew", pady=5)
        self._row(learning_form, 3, "Notes", self.feedback_notes)
        ttk.Button(
            learning_form,
            text="Save Rejection Feedback",
            style="Blue.TButton",
            command=self._save_feedback,
        ).grid(row=4, column=1, sticky="w", pady=(10, 0))

        learning_list_frame = ttk.LabelFrame(
            learning_tab, text="Recent learning", padding=8
        )
        learning_list_frame.pack(fill="both", expand=True)
        columns = ("company", "website", "reason", "notes", "date")
        self.learning_tree = ttk.Treeview(
            learning_list_frame, columns=columns, show="headings"
        )
        for column, title, width in [
            ("company", "Company", 190),
            ("website", "Website", 180),
            ("reason", "Reason", 150),
            ("notes", "Notes", 260),
            ("date", "Date", 150),
        ]:
            self.learning_tree.heading(column, text=title)
            self.learning_tree.column(column, width=width, stretch=True)
        learn_scroll = ttk.Scrollbar(
            learning_list_frame, command=self.learning_tree.yview
        )
        self.learning_tree.configure(yscrollcommand=learn_scroll.set)
        self.learning_tree.pack(side="left", fill="both", expand=True)
        learn_scroll.pack(side="right", fill="y")

        knowledge_header = ttk.Frame(knowledge_tab, style="Card.TFrame")
        knowledge_header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            knowledge_header,
            text="Darwill messaging and value propositions",
            style="Card.TLabel",
            font=("Segoe UI", 12, "bold"),
        ).pack(side="left")
        ttk.Button(
            knowledge_header,
            text="Reset Defaults",
            style="Secondary.TButton",
            command=self._reset_darwill_knowledge,
        ).pack(side="right")
        ttk.Button(
            knowledge_header,
            text="Save Knowledge",
            style="Blue.TButton",
            command=self._save_darwill_knowledge,
        ).pack(side="right", padx=(0, 8))

        knowledge_canvas = tk.Canvas(
            knowledge_tab, bg=WHITE, highlightthickness=0
        )
        knowledge_scroll = ttk.Scrollbar(
            knowledge_tab, orient="vertical", command=knowledge_canvas.yview
        )
        knowledge_content = ttk.Frame(
            knowledge_canvas, padding=8, style="Card.TFrame"
        )
        knowledge_content.bind(
            "<Configure>",
            lambda _e: knowledge_canvas.configure(
                scrollregion=knowledge_canvas.bbox("all")
            ),
        )
        knowledge_canvas.create_window(
            (0, 0), window=knowledge_content, anchor="nw"
        )
        knowledge_canvas.configure(yscrollcommand=knowledge_scroll.set)
        knowledge_canvas.pack(side="left", fill="both", expand=True)
        knowledge_scroll.pack(side="right", fill="y")

        self.knowledge_company_summary = tk.Text(
            knowledge_content, height=4, wrap="word"
        )
        self.knowledge_meeting_cta = tk.Text(
            knowledge_content, height=3, wrap="word"
        )
        ttk.Label(
            knowledge_content, text="Darwill company summary",
            style="Card.TLabel", font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")
        self.knowledge_company_summary.pack(fill="x", pady=(4, 12))
        ttk.Label(
            knowledge_content, text="Default meeting CTA",
            style="Card.TLabel", font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")
        self.knowledge_meeting_cta.pack(fill="x", pady=(4, 12))

        self.knowledge_offering_widgets = {}
        for strategy in STRATEGY_NAMES:
            frame = ttk.LabelFrame(
                knowledge_content, text=strategy, padding=10
            )
            frame.pack(fill="x", pady=(0, 10))
            ttk.Label(
                frame, text="Value proposition", style="Card.TLabel"
            ).pack(anchor="w")
            value_box = tk.Text(frame, height=3, wrap="word")
            value_box.pack(fill="x", pady=(4, 8))
            ttk.Label(
                frame, text="Proof points / talking points",
                style="Card.TLabel",
            ).pack(anchor="w")
            proof_box = tk.Text(frame, height=3, wrap="word")
            proof_box.pack(fill="x", pady=(4, 0))
            self.knowledge_offering_widgets[strategy] = (
                value_box, proof_box
            )
        self._populate_darwill_knowledge()

        footer = tk.Frame(self, bg="#E6EEF5", height=34)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Label(
            footer,
            text=f"Developed by {DEVELOPER_NAME}",
            bg="#E6EEF5",
            fg=MUTED,
            font=("Segoe UI", 8),
        ).pack(side="left", padx=16)
        self.footer_status = tk.StringVar(value="Ready")
        tk.Label(
            footer,
            textvariable=self.footer_status,
            bg="#E6EEF5",
            fg=NAVY,
            font=("Segoe UI Semibold", 8),
        ).pack(side="right", padx=16)

    def _formatted_number_row(self, parent, row, label, variable):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=5
        )
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=5)
        def normalize(_event=None):
            parsed = parse_number(variable.get())
            if parsed is not None:
                variable.set(f"{parsed:,}")
        entry.bind("<FocusOut>", normalize)
        entry.bind("<Return>", normalize)
        parent.columnconfigure(1, weight=1)

    def _row(self, parent, row, label, variable, show=""):
        ttk.Label(parent, text=label, style="Card.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=5
        )
        ttk.Entry(parent, textvariable=variable, show=show).grid(
            row=row, column=1, sticky="ew", pady=5
        )
        parent.columnconfigure(1, weight=1)

    def _append(self, message):
        self.log.configure(state="normal")
        self.log.insert("end", str(message).rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

        self._dashboard_from_log(str(message))

    def _dashboard_from_log(self, message: str):
        """Presentation-only parsing of existing engine log messages."""
        if not hasattr(self, "preview_vars"):
            return
        line = (message or "").strip()
        lower = line.lower()

        returned = re.search(
            r"zoominfo returned\s+(\d+)\s+companies",
            lower,
        )
        if returned:
            self._update_funnel("zoominfo", int(returned.group(1)), add=True)

        if "tavily" in lower and (
            "research" in lower
            or "verification" in lower
        ):
            self.preview_vars["phase"].set("TAVILY RESEARCH")

        if lower.startswith("qualified "):
            match = re.search(
                r"qualified\s+\d+/\d+:\s+(.+?)\s+\(([\d.]+)\)",
                line,
                re.I,
            )
            if match:
                self.preview_vars["company"].set(match.group(1))
                self.preview_vars["score"].set(match.group(2))
                self.preview_vars["phase"].set("QUALIFIED")
                self.preview_vars["detail"].set(
                    "Accepted into the qualified list. "
                    "Compass is now ranking decision-makers and recovering "
                    "the best available contact data."
                )

        if (
            "skipped" in lower
            or lower.startswith("rejected")
            or "outside target geography" in lower
        ):
            self.preview_vars["phase"].set("FILTERED")
            if (
                lower.startswith("rejected")
                or " skipped " in lower
                or lower.startswith("master db skipped")
                or lower.startswith("hubspot csv index skipped")
            ):
                self._update_funnel("rejected", 1, add=True)

    def _update_funnel(
        self,
        key: str,
        value: int,
        *,
        add: bool = False,
    ):
        if not hasattr(self, "funnel_counts") or key not in self.funnel_counts:
            return
        if add:
            self.funnel_counts[key] += int(value)
        else:
            self.funnel_counts[key] = int(value)
        self.funnel_vars[key].set(f"{self.funnel_counts[key]:,}")

    def _set_prospect_preview(self, payload: dict[str, Any]):
        if not hasattr(self, "preview_vars"):
            return
        for key in self.preview_vars:
            if key in payload and payload[key] not in (None, ""):
                self.preview_vars[key].set(str(payload[key]))


    def _load_settings(self):
        if SETTINGS_FILE.exists():
            try:
                data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
                self.states.set(data.get("states", DEFAULT_STATES))
                self.trades.set(data.get("trades", DEFAULT_TRADES))
                self.search_mode.set(data.get("search_mode", "states"))
                self.custom_trade_keywords.set(data.get("custom_trade_keywords", ""))
                selected_trades = set(data.get(
                    "selected_trades",
                    ["HVAC", "Plumbing", "Electrical", "Pest Control", "Pool Service", "Garage Door"],
                ))
                for name, variable in self.trade_vars.items():
                    variable.set(name in selected_trades)
                self.minimum_revenue.set(format_integer_with_commas(data.get("minimum_revenue", 10_000_000)))
                self.maximum_revenue.set(format_integer_with_commas(data.get("maximum_revenue", 250_000_000)))
                self.target_count.set(data.get("target_count", 20))
                self.candidate_limit.set(data.get("candidate_limit", 100))
                self.contacts_per_company.set(data.get("contacts_per_company", 3))
                self.minimum_fit_score.set(data.get("minimum_fit_score", 50))
                self.naics_codes.set(data.get("naics_codes", "238220,238210,561710,238290"))
                self.employee_min.set(data.get("employee_min", 20))
                self.employee_max.set(data.get("employee_max", 1000))
                self.exclude_crm.set(data.get("exclude_crm", True))
                self.enrich_final_contacts.set(data.get("enrich_final_contacts", False))
                self.smart_email_enrichment.set(data.get("smart_email_enrichment", True))
                self.use_zoominfo_ai_research.set(data.get("use_zoominfo_ai_research", False))
                self.skip_history.set(data.get("skip_history", True))
                self.resume_checkpoint.set(data.get("resume_checkpoint", True))
                self.output_folder.set(data.get("output_folder", str(OUTPUT_DIR)))
                self.territory_zip.set(data.get("territory_zip", ""))
                self.territory_radius.set(data.get("territory_radius", "50"))
                self.research_workers.set(data.get("research_workers", 4))
                self.profile_name.set(data.get("profile_name", "Default"))
                self.show_advanced.set(data.get("show_advanced", False))
                self.hubspot_sender_email.set(data.get("hubspot_sender_email", ""))
                self.use_master_dedup.set(data.get("use_master_dedup", True))
                self.rejected_rereview_days.set(
                    data.get("rejected_rereview_days", 180)
                )
                self.qualified_rereview_days.set(
                    data.get("qualified_rereview_days", 0)
                )
                self.discover_public_contacts.set(
                    data.get("discover_public_contacts", True)
                )
                self.predict_public_emails.set(
                    data.get("predict_public_emails", True)
                )
                self.use_hubspot_csv_index.set(
                    data.get("use_hubspot_csv_index", True)
                )
                self.deep_contact_recovery.set(
                    data.get("deep_contact_recovery", True)
                )
                self.deep_recovery_contact_limit.set(
                    data.get("deep_recovery_contact_limit", 3)
                )
                self.block_unverified_sequence_emails.set(
                    data.get("block_unverified_sequence_emails", True)
                )
                self.master_csv_path.set(
                    data.get("master_csv_path", "")
                )
                self.master_csv_auto_sync.set(
                    data.get("master_csv_auto_sync", True)
                )
            except Exception:
                pass
        try:
            self.client_id.set(keyring.get_password(SERVICE, "client_id") or "")
            self.client_secret.set(keyring.get_password(SERVICE, "client_secret") or "")
            self.tavily_key.set(keyring.get_password(SERVICE, "tavily_api_key") or "")
            self.hubspot_token.set(keyring.get_password(SERVICE, "hubspot_token") or "")
        except Exception:
            pass

    def _save_credentials(self):
        if self.client_id.get():
            keyring.set_password(SERVICE, "client_id", self.client_id.get())
        if self.client_secret.get():
            keyring.set_password(SERVICE, "client_secret", self.client_secret.get())
        if self.tavily_key.get():
            keyring.set_password(SERVICE, "tavily_api_key", self.tavily_key.get())
        if self.hubspot_token.get():
            keyring.set_password(SERVICE, "hubspot_token", self.hubspot_token.get())
        self._refresh_connection_status()
        messagebox.showinfo(APP_TITLE, "Credentials saved securely in Windows Credential Manager.")

    def _save_settings(self):
        SETTINGS_FILE.write_text(json.dumps({
            "states": self.states.get(),
            "trades": self.trades.get(),
            "search_mode": self.search_mode.get(),
            "selected_trades": [
                name for name, variable in self.trade_vars.items() if variable.get()
            ],
            "custom_trade_keywords": self.custom_trade_keywords.get(),
            "minimum_revenue": parse_number(self.minimum_revenue.get()) or 0,
            "maximum_revenue": parse_number(self.maximum_revenue.get()) or 0,
            "target_count": self.target_count.get(),
            "candidate_limit": self.candidate_limit.get(),
            "contacts_per_company": self.contacts_per_company.get(),
            "minimum_fit_score": self.minimum_fit_score.get(),
            "naics_codes": self.naics_codes.get(),
            "employee_min": self.employee_min.get(),
            "employee_max": self.employee_max.get(),
            "exclude_crm": self.exclude_crm.get(),
            "enrich_final_contacts": self.enrich_final_contacts.get(),
            "smart_email_enrichment": self.smart_email_enrichment.get(),
            "use_zoominfo_ai_research": self.use_zoominfo_ai_research.get(),
            "skip_history": self.skip_history.get(),
            "resume_checkpoint": self.resume_checkpoint.get(),
            "output_folder": self.output_folder.get(),
            "territory_zip": self.territory_zip.get(),
            "territory_radius": self.territory_radius.get(),
            "research_workers": self.research_workers.get(),
            "profile_name": self.profile_name.get(),
            "show_advanced": self.show_advanced.get(),
            "hubspot_sender_email": self.hubspot_sender_email.get(),
            "use_master_dedup": self.use_master_dedup.get(),
            "rejected_rereview_days": self.rejected_rereview_days.get(),
            "qualified_rereview_days": self.qualified_rereview_days.get(),
            "discover_public_contacts": self.discover_public_contacts.get(),
            "predict_public_emails": self.predict_public_emails.get(),
            "use_hubspot_csv_index": self.use_hubspot_csv_index.get(),
            "deep_contact_recovery": self.deep_contact_recovery.get(),
            "deep_recovery_contact_limit": self.deep_recovery_contact_limit.get(),
            "block_unverified_sequence_emails": self.block_unverified_sequence_emails.get(),
            "master_csv_path": self.master_csv_path.get(),
            "master_csv_auto_sync": self.master_csv_auto_sync.get(),
        }, indent=2), encoding="utf-8")
        self._save_credentials()

    def _refresh_hubspot_index(self):
        if not hasattr(self, "hubspot_index_tree"):
            return
        db = HistoryDB(DB_PATH)
        stats = db.hubspot_index_stats()
        self.hubspot_index_company_count.set(
            f"{stats['company_count']:,}"
        )
        self.hubspot_index_contact_count.set(
            f"{stats['contact_count']:,}"
        )
        last_import = stats.get("last_import")
        if last_import:
            self.hubspot_index_last_import.set(
                f"{last_import[0][:10]} • {last_import[2]}"
            )
        else:
            self.hubspot_index_last_import.set("No imports yet")

        self.hubspot_index_tree.delete(
            *self.hubspot_index_tree.get_children()
        )
        for row in db.hubspot_import_history():
            self.hubspot_index_tree.insert(
                "", "end",
                values=(
                    row[0], row[1], row[2], row[3],
                    row[4], row[5], row[6], row[7],
                ),
            )

    def _import_hubspot_csv(self, expected_kind: str, mode: str):
        path_value = filedialog.askopenfilename(
            title=(
                f"Select HubSpot {expected_kind.title()} CSV"
            ),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path_value:
            return
        path = Path(path_value)
        try:
            with path.open(
                "r", encoding="utf-8-sig", newline="",
                errors="replace"
            ) as handle:
                reader = csv.reader(handle)
                fieldnames = next(reader, [])
            detected = hubspot_csv_kind(fieldnames)
            if detected != expected_kind:
                if not messagebox.askyesno(
                    APP_TITLE,
                    (
                        f"This file appears to be a {detected or 'different'} "
                        f"export, not {expected_kind}. Import it anyway?"
                    ),
                    icon="warning",
                ):
                    return

            if mode == "replace":
                if not messagebox.askyesno(
                    APP_TITLE,
                    (
                        f"Replace the current HubSpot {expected_kind} index "
                        f"with {path.name}?\n\n"
                        "The Master Prospect Database is not affected."
                    ),
                    icon="warning",
                ):
                    return

            result = HistoryDB(DB_PATH).import_hubspot_csv(
                path, mode=mode
            )
            self._refresh_hubspot_index()
            messagebox.showinfo(
                APP_TITLE,
                (
                    f"HubSpot {result['kind']} index updated.\n\n"
                    f"Rows read: {result['rows_read']:,}\n"
                    f"New records: {result['rows_indexed']:,}\n"
                    f"Updated records: {result['rows_updated']:,}\n"
                    f"Skipped rows: {result['rows_skipped']:,}"
                ),
            )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"HubSpot CSV import failed:\n\n{exc}",
            )

    def _clear_hubspot_index(self):
        if not messagebox.askyesno(
            APP_TITLE,
            (
                "Clear all imported HubSpot company and contact index data?\n\n"
                "This does not delete the Master Prospect Database or HubSpot records."
            ),
            icon="warning",
        ):
            return
        HistoryDB(DB_PATH).clear_hubspot_index("all")
        self._refresh_hubspot_index()

    def _resolve_master_csv_path(self) -> Path | None:
        value = self.master_csv_path.get().strip()
        if value:
            return Path(value).expanduser()
        candidates = detect_master_csv_candidates()
        if candidates:
            selected = candidates[0]
            self.master_csv_path.set(str(selected))
            self._save_settings()
            return selected
        return None

    def _browse_master_csv(self):
        current = self.master_csv_path.get().strip()
        initial_dir = (
            str(Path(current).parent)
            if current
            else str(Path.home() / "Documents")
        )
        selected = filedialog.askopenfilename(
            title="Select Permanent Master Prospect Database CSV",
            initialdir=initial_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not selected:
            return
        self.master_csv_path.set(selected)
        self._save_settings()
        self._load_master_csv()

    def _load_master_csv(self, show_message: bool = True):
        path = self._resolve_master_csv_path()
        if path is None:
            selected = filedialog.askopenfilename(
                title="Select Permanent Master Prospect Database CSV",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
            if not selected:
                return
            path = Path(selected)
            self.master_csv_path.set(str(path))
            self._save_settings()

        if not path.exists():
            if not messagebox.askyesno(
                APP_TITLE,
                (
                    f"The selected master CSV does not exist:\n\n{path}\n\n"
                    "Create it from the application's current internal "
                    "master database?"
                ),
            ):
                return
            result = sync_internal_master_to_csv(
                HistoryDB(DB_PATH),
                path,
                create_backup=False,
            )
            self._refresh_master_csv_status()
            if show_message:
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        "Permanent master CSV created.\n\n"
                        f"Records written: {result['records']:,}\n"
                        f"File: {path}"
                    ),
                )
            return

        try:
            result = import_master_csv_into_internal_db(
                HistoryDB(DB_PATH),
                path,
            )
            self._refresh_master_database()
            self._refresh_master_csv_status()
            self._save_settings()
            if show_message:
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        "Permanent master CSV loaded.\n\n"
                        f"Rows read: {result['rows_read']:,}\n"
                        f"New internal records: {result['imported']:,}\n"
                        f"Updated internal records: {result['updated']:,}\n"
                        f"Skipped: {result['skipped']:,}\n"
                        f"Internal master total: "
                        f"{result['total_internal']:,}"
                    ),
                )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"Could not load the permanent master CSV:\n\n{exc}",
            )

    def _sync_master_csv_now(self, show_message: bool = True):
        path = self._resolve_master_csv_path()
        if path is None:
            selected = filedialog.asksaveasfilename(
                title="Create Permanent Master Prospect Database CSV",
                defaultextension=".csv",
                initialfile=DEFAULT_MASTER_CSV_NAME,
                filetypes=[("CSV files", "*.csv")],
            )
            if not selected:
                return
            path = Path(selected)
            self.master_csv_path.set(str(path))
            self._save_settings()

        try:
            result = sync_internal_master_to_csv(
                HistoryDB(DB_PATH),
                path,
                create_backup=True,
            )
            self._refresh_master_csv_status()
            if show_message:
                message = (
                    "Permanent master CSV synchronized.\n\n"
                    f"Records: {result['records']:,}\n"
                    f"File: {result['path']}"
                )
                if result["backup"]:
                    message += (
                        "\n\nBackup created:\n"
                        f"{result['backup']}"
                    )
                messagebox.showinfo(APP_TITLE, message)
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"Could not synchronize the permanent master CSV:\n\n{exc}",
            )

    def _open_master_csv_folder(self):
        path = self._resolve_master_csv_path()
        if path is None:
            messagebox.showinfo(
                APP_TITLE,
                "Select or create a permanent master CSV first.",
            )
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        os.startfile(str(path.parent))

    def _refresh_master_csv_status(self):
        path = self._resolve_master_csv_path()
        total = HistoryDB(DB_PATH).master_counts().get("Total", 0)
        self.master_csv_company_count.set(f"{total:,}")
        if path is None:
            self.master_csv_loaded_file.set(
                "No permanent master CSV selected"
            )
            self.master_csv_last_updated.set("Never")
            return

        self.master_csv_loaded_file.set(f"Loaded: {path}")
        if path.exists():
            modified = datetime.fromtimestamp(
                path.stat().st_mtime
            ).strftime("%Y-%m-%d %I:%M %p")
            self.master_csv_last_updated.set(modified)
        else:
            self.master_csv_last_updated.set(
                "Selected file has not been created"
            )

    def _initialize_permanent_master_csv(self):
        path = self._resolve_master_csv_path()
        if path is None:
            self._refresh_master_csv_status()
            return
        if path.exists():
            try:
                import_master_csv_into_internal_db(
                    HistoryDB(DB_PATH),
                    path,
                )
            except Exception:
                pass
        self._refresh_master_database()
        self._refresh_master_csv_status()

    def _refresh_master_database(self):
        if not hasattr(self, "master_tree"):
            return
        db = HistoryDB(DB_PATH)
        rows = db.master_prospects(
            self.master_status_filter.get(),
            self.master_search.get(),
        )
        self.master_tree.delete(*self.master_tree.get_children())
        for row in rows:
            (
                key, company, domain, city, state, trade, revenue,
                employees, status, decision, fit, reviewed, next_review,
                hubspot_id, sequence_status, reason,
            ) = row
            self.master_tree.insert(
                "",
                "end",
                iid=key,
                values=(
                    company or "",
                    domain or "",
                    state or "",
                    trade or "",
                    f"${revenue:,.0f}" if revenue else "",
                    f"{employees:,}" if employees else "",
                    status or "",
                    decision or "",
                    f"{fit:.0f}" if fit is not None else "",
                    (reviewed or "")[:10],
                    (next_review or "")[:10],
                    hubspot_id or "",
                    sequence_status or "",
                ),
                tags=(reason or "",),
            )
        counts = db.master_counts()
        for name, variable in self.master_count_vars.items():
            variable.set(str(counts.get(name, 0)))

    def _load_master_detail(self):
        if not hasattr(self, "master_tree"):
            return
        selected = self.master_tree.selection()
        if not selected:
            return
        key = selected[0]
        row = HistoryDB(DB_PATH).conn.execute(
            """SELECT company_name, lifecycle_status, notes
               FROM master_prospects WHERE prospect_key=?""",
            (key,),
        ).fetchone()
        if not row:
            return
        self.master_selected_key.set(key)
        self.master_selected_company.set(row[0] or "")
        self.master_selected_status.set(row[1] or "New")
        self.master_selected_notes.set(row[2] or "")

    def _update_master_status(self):
        key = self.master_selected_key.get()
        if not key:
            messagebox.showinfo(
                APP_TITLE, "Select a company in the master database."
            )
            return
        db = HistoryDB(DB_PATH)
        record = db.conn.execute(
            """SELECT company_id, website, company_name
               FROM master_prospects WHERE prospect_key=?""",
            (key,),
        ).fetchone()
        if not record:
            return
        db.update_master_lifecycle(
            record[0] or "",
            record[1] or "",
            record[2] or "",
            self.master_selected_status.get(),
            notes=self.master_selected_notes.get(),
        )
        self._refresh_master_database()
        if (
            self.master_csv_auto_sync.get()
            and self.master_csv_path.get().strip()
        ):
            self._sync_master_csv_now(show_message=False)

    def _export_master_database(self):
        path = OUTPUT_DIR / (
            "Darwill_Master_Prospect_Database_"
            + datetime.now().strftime("%Y-%m-%d_%H%M")
            + ".csv"
        )
        rows = HistoryDB(DB_PATH).conn.execute(
            """SELECT company_name, domain, website, city, state, trade,
                      revenue, employees, industry, lifecycle_status,
                      qualification_decision, qualification_reason, fit_score,
                      first_seen_at, last_reviewed_at, next_review_at,
                      approved_at, hubspot_company_id, sequence_status,
                      meeting_status, customer_status, source_profile, notes
               FROM master_prospects ORDER BY company_name"""
        ).fetchall()
        headers = [
            "Company", "Domain", "Website", "City", "State", "Trade",
            "Revenue", "Employees", "Industry", "Lifecycle Status",
            "Qualification Decision", "Qualification Reason", "Fit Score",
            "First Seen", "Last Reviewed", "Next Review", "Approved At",
            "HubSpot Company ID", "Sequence Status", "Meeting Status",
            "Customer Status", "Source Profile", "Notes",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(headers)
            writer.writerows(rows)
        os.startfile(str(OUTPUT_DIR))
        messagebox.showinfo(
            APP_TITLE,
            f"Master prospect database exported:\n\n{path}",
        )

    def _save_deliverability_settings(self):
        self.deliverability_rules["minimum_enrollment_score"] = int(
            self.minimum_inbox_score.get()
        )
        self.deliverability_rules["sender_authentication"] = {
            "spf_confirmed": bool(self.spf_confirmed.get()),
            "dkim_confirmed": bool(self.dkim_confirmed.get()),
            "dmarc_confirmed": bool(self.dmarc_confirmed.get()),
        }
        DELIVERABILITY_RULES_FILE.write_text(
            json.dumps(self.deliverability_rules, indent=2),
            encoding="utf-8",
        )

    def _delivery_analysis_for_item(
        self, item: ReviewQueueItem
    ) -> dict[str, Any]:
        self._save_deliverability_settings()
        return analyze_deliverability(
            item.subject_line,
            item.email_body,
            item.company_name,
            item.contact_first_name,
            item.recommended_strategy,
            self.deliverability_rules,
        )

    def _update_item_deliverability(
        self,
        item: ReviewQueueItem,
    ) -> dict[str, Any]:
        analysis = self._delivery_analysis_for_item(item)
        item.inbox_readiness_score = analysis["inbox_readiness_score"]
        item.subject_quality_score = analysis["subject_quality_score"]
        item.personalization_score = analysis["personalization_score"]
        item.content_risk_score = analysis["content_risk_score"]
        item.human_tone_score = analysis["human_tone_score"]
        item.formatting_score = analysis["formatting_score"]
        item.spam_risk = analysis["spam_risk"]
        item.send_ready = analysis["send_ready"]
        item.deliverability_issues = " | ".join(analysis["issues"])
        item.deliverability_recommendations = " | ".join(
            analysis["recommendations"]
        )
        item.updated_at = now_iso()
        return analysis

    def _populate_delivery_item(
        self,
        item: ReviewQueueItem,
        analysis: dict[str, Any] | None = None,
    ):
        if not hasattr(self, "delivery_overall"):
            return
        if analysis is None:
            analysis = self._update_item_deliverability(item)

        self.delivery_company.set(item.company_name)
        self.delivery_contact.set(
            f"{item.contact_name} — {item.contact_title}"
        )
        self.delivery_overall.set(
            f"{analysis['inbox_readiness_score']} / 100"
        )
        readiness = (
            "Ready for review"
            if analysis["send_ready"]
            else "Needs changes"
        )
        self.delivery_risk.set(
            f"{analysis['spam_risk']} heuristic risk • {readiness}"
        )
        self.delivery_subject_score.set(
            f"{analysis['subject_quality_score']}%"
        )
        self.delivery_personalization_score.set(
            f"{analysis['personalization_score']}%"
        )
        self.delivery_content_score.set(
            f"{analysis['content_risk_score']}%"
        )
        self.delivery_tone_score.set(
            f"{analysis['human_tone_score']}%"
        )
        self.delivery_format_score.set(
            f"{analysis['formatting_score']}%"
        )
        self.delivery_auth_score.set(
            f"{analysis['authentication_score']}%"
        )
        self.delivery_word_count.set(str(analysis["word_count"]))
        self.delivery_link_count.set(str(analysis["link_count"]))
        self.delivery_cta_count.set(str(analysis["cta_count"]))
        self.delivery_outcome.set(item.outcome or "Not Sent")

        self.delivery_issues.delete("1.0", "end")
        issue_lines = analysis["issues"] or [
            "No material content risks were identified by the heuristic analyzer."
        ]
        self.delivery_issues.insert(
            "1.0",
            "\n".join(f"• {line}" for line in issue_lines),
        )

        self.delivery_recommendations.delete("1.0", "end")
        rec_lines = analysis["recommendations"] or [
            "The draft is concise, personalized, and ready for human review."
        ]
        positives = analysis.get("positives", [])
        output = [f"• {line}" for line in rec_lines]
        if positives:
            output += ["", "Working well:"] + [
                f"✓ {line}" for line in positives
            ]
        self.delivery_recommendations.insert(
            "1.0", "\n".join(output)
        )

        alternatives = analysis["subject_alternatives"]
        variables = [
            self.delivery_alt_subject_1,
            self.delivery_alt_subject_2,
            self.delivery_alt_subject_3,
        ]
        for variable, subject in zip(variables, alternatives):
            variable.set(subject)

    def _analyze_selected_delivery(self):
        item = self._queue_item_by_id(self.review_queue_id.get())
        if not item:
            messagebox.showinfo(
                APP_TITLE,
                "Select a Deal Desk item first.",
            )
            return
        # Capture unsaved Deal Desk edits before analysis.
        item.subject_line = self.review_subject.get().strip()
        item.email_body = self.review_email_body.get(
            "1.0", "end"
        ).strip()
        analysis = self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)
        self._populate_delivery_item(item, analysis)
        self._refresh_deal_desk()

    def _apply_subject_alternative(self, index: int):
        item = self._queue_item_by_id(self.review_queue_id.get())
        if not item:
            messagebox.showinfo(
                APP_TITLE, "Select a Deal Desk item first."
            )
            return
        variables = [
            self.delivery_alt_subject_1,
            self.delivery_alt_subject_2,
            self.delivery_alt_subject_3,
        ]
        if not 1 <= index <= len(variables):
            return
        subject = variables[index - 1].get().strip()
        item.subject_line = subject
        item.subject_variant = chr(64 + index)
        self.review_subject.set(subject)
        analysis = self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)
        self._populate_delivery_item(item, analysis)
        self._refresh_deal_desk()

    def _apply_safer_rewrite(self):
        item = self._queue_item_by_id(self.review_queue_id.get())
        if not item:
            messagebox.showinfo(
                APP_TITLE, "Select a Deal Desk item first."
            )
            return
        rewritten = safer_email_rewrite(
            item.email_body,
            item.company_name,
            item.contact_first_name,
            item.recommended_strategy,
        )
        item.email_body = rewritten
        self.review_email_body.delete("1.0", "end")
        self.review_email_body.insert("1.0", rewritten)
        analysis = self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)
        self._populate_delivery_item(item, analysis)
        self._refresh_deal_desk()

    def _save_delivery_to_queue(self):
        item = self._queue_item_by_id(self.review_queue_id.get())
        if not item:
            messagebox.showinfo(
                APP_TITLE, "Select a Deal Desk item first."
            )
            return
        item.subject_line = self.review_subject.get().strip()
        item.email_body = self.review_email_body.get(
            "1.0", "end"
        ).strip()
        analysis = self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)
        self._populate_delivery_item(item, analysis)
        self._refresh_deal_desk()
        messagebox.showinfo(
            APP_TITLE,
            "Deliverability analysis and edits were saved.",
        )

    def _save_delivery_outcome(self):
        item = self._queue_item_by_id(self.review_queue_id.get())
        if not item:
            messagebox.showinfo(
                APP_TITLE, "Select a Deal Desk item first."
            )
            return
        item.outcome = self.delivery_outcome.get()
        item.updated_at = now_iso()
        lifecycle_map = {
            "Meeting Booked": "Meeting",
            "Replied": "In Sequence",
            "Sent": "In Sequence",
            "Opened": "In Sequence",
            "No Response": "In Sequence",
            "Spam/Junk": "In Sequence",
            "Bounced": "In Sequence",
        }
        mapped_status = lifecycle_map.get(item.outcome)
        if mapped_status:
            HistoryDB(DB_PATH).update_master_lifecycle(
                item.company_id,
                item.company_website,
                item.company_name,
                mapped_status,
                sequence_status=item.sequence_name,
                meeting_status=(
                    "Meeting Booked"
                    if item.outcome == "Meeting Booked"
                    else ""
                ),
            )
        save_outreach_queue(self.review_queue)
        self._refresh_master_database()
        messagebox.showinfo(
            APP_TITLE,
            f"Outcome saved: {item.outcome}",
        )

    def _show_contact_intelligence_tab(self):
        if not self.review_queue_id.get():
            messagebox.showinfo(
                APP_TITLE,
                "Select a contact in the Deal Desk first.",
            )
            return
        self.review_notebook.select(self.intelligence_tab)

    def _show_contact_acquisition_tab(self):
        if not self.review_queue_id.get():
            messagebox.showinfo(
                APP_TITLE,
                "Select a contact in the Deal Desk first.",
            )
            return
        self.review_notebook.select(self.acquisition_tab)

    def _queue_item_by_id(self, queue_id: str) -> ReviewQueueItem | None:
        return next(
            (item for item in self.review_queue if item.queue_id == queue_id),
            None,
        )

    def _set_initial_deal_desk_split(self, pane):
        """Set a practical first-open split without blocking user resizing."""
        try:
            width = pane.winfo_width()
            if width > 400:
                pane.sashpos(0, max(500, int(width * 0.48)))
        except (tk.TclError, IndexError):
            pass

    def _refresh_deal_desk(self):
        if not hasattr(self, "deal_tree"):
            return
        self.review_queue = load_outreach_queue()
        self.deal_tree.delete(*self.deal_tree.get_children())
        selected_filter = self.deal_desk_filter.get()

        counts = {
            "pending": 0,
            "verification": 0,
            "approved": 0,
            "synced": 0,
            "enrolled": 0,
        }
        for item in self.review_queue:
            if item.status == "Pending Review":
                counts["pending"] += 1
            if item.contact_email_verification_status in {
                "Needs Verification",
                "Missing",
            }:
                counts["verification"] += 1
            if item.status == "Approved":
                counts["approved"] += 1
            if item.hubspot_status == "Synced":
                counts["synced"] += 1
            if item.enrollment_status == "Enrolled":
                counts["enrolled"] += 1

            if selected_filter != "All":
                if (
                    selected_filter == "Synced"
                    and item.hubspot_status != "Synced"
                ):
                    continue
                elif (
                    selected_filter == "Enrolled"
                    and item.enrollment_status != "Enrolled"
                ):
                    continue
                elif selected_filter == "Needs Verification":
                    if item.contact_email_verification_status not in {
                        "Needs Verification",
                        "Missing",
                    }:
                        continue
                elif selected_filter not in {
                    "Synced",
                    "Enrolled",
                    "Needs Verification",
                } and item.status != selected_filter:
                    continue

            tag = "pending"
            if item.enrollment_status == "Enrolled":
                tag = "enrolled"
            elif item.hubspot_status == "Synced":
                tag = "synced"
            elif item.status == "Approved":
                tag = "approved"
            elif item.status == "Rejected":
                tag = "rejected"
            elif item.contact_email_verification_status in {
                "Needs Verification",
                "Missing",
            }:
                tag = "verification"

            self.deal_tree.insert(
                "",
                "end",
                iid=item.queue_id,
                values=(
                    (
                        item.contact_outreach_order_label
                        or (
                            f"#{item.contact_outreach_order}"
                            if item.contact_outreach_order
                            else item.contact_rank
                        )
                    ),
                    item.status,
                    item.company_name,
                    f"{item.contact_name} — {item.contact_title}",
                    item.recommended_strategy,
                    f"{item.outreach_confidence}%",
                    (
                        f"{item.inbox_readiness_score}%"
                        if item.inbox_readiness_score
                        else "—"
                    ),
                    item.hubspot_status,
                    item.enrollment_status,
                ),
                tags=(tag,),
            )

        if hasattr(self, "deal_kpi_vars"):
            for key, value in counts.items():
                self.deal_kpi_vars[key].set(f"{value:,}")

    def _load_selected_queue_item(self):
        if not hasattr(self, "deal_tree"):
            return
        selected = self.deal_tree.selection()
        if not selected:
            return
        item = self._queue_item_by_id(selected[0])
        if not item:
            return
        self.review_queue_id.set(item.queue_id)
        self.review_company.set(item.company_name)

        order = item.contact_outreach_order or (
            1 if item.contact_rank == "Primary"
            else 2 if item.contact_rank == "Secondary"
            else 3 if item.contact_rank in {"Third", "Fallback"}
            else 0
        )
        order_label = item.contact_outreach_order_label or (
            "PRIMARY — CONTACT FIRST" if order == 1
            else "SECONDARY — CONTACT SECOND" if order == 2
            else "THIRD — CONTACT THIRD" if order == 3
            else item.contact_rank.upper() if item.contact_rank
            else "OUTREACH ORDER NOT AVAILABLE"
        )
        item.contact_outreach_order = order
        item.contact_outreach_order_label = order_label

        self.review_outreach_order.set(
            f"#{order}  {order_label}" if order else order_label
        )
        self.review_contact_source.set(
            f"Found through: "
            f"{item.contact_source_type or 'Legacy queue record'}"
        )
        self.review_contact.set(
            f"{item.contact_name} — {item.contact_title}"
        )
        self.review_strategy.set(
            f"Recommended strategy: {item.recommended_strategy}"
        )
        self.review_confidence.set(
            f"Review status: {item.status}"
        )
        self.review_subject.set(item.subject_line)
        self.review_notes.set(item.reviewer_notes)
        self.review_email_body.delete("1.0", "end")
        self.review_email_body.insert("1.0", item.email_body)
        decision_confidence = (
            item.contact_decision_confidence
            or min(99, max(1, int(item.outreach_confidence or 0)))
        )
        source_type = item.contact_source_type or (
            "ZoomInfo / legacy queue record"
            if item.contact_id
            and not str(item.contact_id).startswith("public-")
            else "Public research / legacy queue record"
        )
        email_status = item.contact_email_status or (
            "Available — source not recorded"
            if item.contact_email
            else "Missing"
        )
        phone_status = item.contact_phone_status or (
            "Available — source not recorded"
            if item.contact_phone
            else "Missing"
        )
        recommendation_reason = (
            item.contact_recommendation_reason
            or item.why_contact
            or "Selected by the contact-ranking model."
        )
        research_summary = (
            item.contact_research_summary
            or item.evidence
            or "No additional contact-specific research was stored."
        )
        source_urls = (
            item.contact_research_sources
            or item.contact_source_url
            or item.sources
            or "No source URL was stored."
        )

        self.review_metric_vars["outreach"].set(
            f"{item.outreach_confidence or 0}%"
        )
        self.review_metric_vars["decision"].set(
            f"{decision_confidence}%"
        )
        self.review_metric_vars["email"].set(
            (
                f"{item.contact_email_confidence}%"
                if item.contact_email
                else "Missing"
            )
        )
        self.review_metric_vars["phone"].set(
            (
                f"{item.contact_phone_confidence}%"
                if item.contact_phone
                else (
                    "Company #"
                    if item.contact_public_company_phone
                    else "Missing"
                )
            )
        )
        self.review_metric_vars["inbox"].set(
            (
                f"{item.inbox_readiness_score}%"
                if item.inbox_readiness_score
                else "—"
            )
        )
        self.review_company_detail.set(
            " · ".join(
                part
                for part in [
                    (
                        f"{item.company_city}, {item.company_state}"
                        if item.company_city and item.company_state
                        else item.company_state
                    ),
                    (
                        f"${item.company_revenue:,.0f} revenue"
                        if item.company_revenue
                        else ""
                    ),
                    (
                        f"{item.company_employees:,} employees"
                        if item.company_employees
                        else ""
                    ),
                ]
                if part
            )
            or "Company size and location were not stored."
        )
        company_brief = build_company_brief(item)
        company_presentation = build_company_presentation(item)
        if hasattr(self, "company_intelligence_vars"):
            for key, value in company_presentation.items():
                if key in self.company_intelligence_vars:
                    self.company_intelligence_vars[key].set(str(value))
        if hasattr(self, "company_intelligence_text"):
            self.company_intelligence_text.configure(state="normal")
            self.company_intelligence_text.delete("1.0", "end")
            self.company_intelligence_text.insert(
                "1.0",
                company_presentation.get("evidence_text", ""),
            )
            self.company_intelligence_text.configure(state="disabled")

        executive_recommendation = (
            recommendation_reason
            if recommendation_reason
            and recommendation_reason
            != "Contact recommendation will appear here."
            else (
                f"{company_brief.fit_summary} "
                f"Recommended angle: {company_brief.recommended_angle}"
            )
        )
        self.review_recommendation.set(executive_recommendation)
        self.review_email_detail.set(
            (
                f"{item.contact_email or 'No email found'} · "
                f"{email_status} · "
                f"{item.contact_email_confidence or 0}% confidence · "
                f"{item.contact_email_verification_status or 'Not verified'}"
            )
        )
        self.review_phone_detail.set(
            (
                f"{item.contact_phone or item.contact_public_company_phone or 'No phone found'} · "
                f"{phone_status} · "
                f"{item.contact_phone_confidence or 0}% confidence"
            )
        )
        if hasattr(self, "contact_intelligence_metric_vars"):
            self.contact_intelligence_metric_vars["rank"].set(
                f"#{order}" if order else item.contact_rank or "—"
            )
            self.contact_intelligence_metric_vars["decision"].set(
                f"{decision_confidence}%"
            )
            self.contact_intelligence_metric_vars["email"].set(
                f"{item.contact_email_confidence or 0}%"
            )
            self.contact_intelligence_metric_vars["phone"].set(
                f"{item.contact_phone_confidence or 0}%"
            )

        intelligence = (
            "WHY THIS PERSON RANKED HERE\n"
            f"• {recommendation_reason}\n"
            f"• Outreach order: #{order if order else '?'} — {order_label}\n"
            f"• Decision-maker confidence: {decision_confidence}%\n"
            f"• Discovery source: {source_type}\n\n"
            "CONTACTABILITY\n"
            f"• Email: {item.contact_email or 'Not found'}\n"
            f"• Email status: {email_status}\n"
            f"• Verification: "
            f"{item.contact_email_verification_status or 'Unknown'}\n"
            f"• Recovery method: "
            f"{item.contact_email_recovery_method or 'Not used'}\n"
            f"• Phone: "
            f"{item.contact_phone or item.contact_public_company_phone or 'Not found'}\n"
            f"• Phone status: {phone_status}\n\n"
            "LIVE RESEARCH SUMMARY\n"
            f"{research_summary}\n\n"
            "SOURCE EVIDENCE\n"
            f"{source_urls}\n\n"
            "DATA STATUS\n"
            f"{item.contact_data_status or 'No detailed status recorded.'}"
        )
        self.contact_intelligence_text.delete("1.0", "end")
        self.contact_intelligence_text.insert("1.0", intelligence)
        self.contact_acquisition_text.delete("1.0", "end")
        email_resolution = build_email_resolution(item)
        if hasattr(self, "company_intelligence_vars"):
            self.company_intelligence_vars["fit"].set(
                company_brief.fit_summary
            )
            self.company_intelligence_vars["confidence"].set(
                f"{company_brief.confidence}%"
            )
            self.company_intelligence_vars["email_status"].set(
                email_resolution.label
            )
            self.company_intelligence_vars["next_action"].set(
                email_resolution.recommended_action
            )
            executive_summary_text = (
                f"{company_brief.company_name} "
                f"{company_brief.fit_summary.lower()} "
                f"{company_brief.residential_assessment} "
                f"{company_brief.scale_assessment} "
                f"{company_brief.growth_assessment} "
                f"Primary opportunity: "
                f"{company_brief.opportunity_assessment} "
                f"Recommended approach: "
                f"{company_brief.recommended_angle} "
                f"Main risk: {company_brief.risk_assessment}"
            )
            self.company_intelligence_vars["executive_summary"].set(
                executive_summary_text
            )

        email_intelligence = build_email_intelligence(item)
        if hasattr(self, "email_intelligence_vars"):
            self.email_intelligence_vars["public_status"].set(
                email_intelligence.public_status
            )
            self.email_intelligence_vars["pattern"].set(
                email_intelligence.pattern_label
            )
            self.email_intelligence_vars["confidence"].set(
                f"{email_intelligence.confidence}%"
            )
            self.email_intelligence_vars["zoominfo"].set(
                email_intelligence.zoominfo_status
            )
            self.email_intelligence_vars["recommendation"].set(
                email_intelligence.recommendation
            )

        try:
            prospect_key = self.intelligence_store.upsert_queue_item(item)
            self.intelligence_store.record_standard_evidence(
                prospect_key, item, email_intelligence
            )
            self.intelligence_store.add_evidence(
                prospect_key,
                "company_intelligence",
                "fit_summary",
                company_brief.fit_summary,
                confidence=company_brief.confidence,
                verification_status="computed",
            )
            self.intelligence_store.add_evidence(
                prospect_key,
                "email_resolution",
                email_resolution.status,
                email_resolution.explanation,
                contact_key=self.intelligence_store.contact_key(item),
                confidence=email_resolution.confidence,
                verification_status="computed",
            )
            self.intelligence_store.append_timeline(
                prospect_key,
                "reviewed_in_deal_desk",
                "Deal Desk record reviewed",
                "ui",
                {
                    "queue_id": item.queue_id,
                    "contact_id": item.contact_id,
                    "status": item.status,
                },
            )
        except Exception:
            pass

        acquisition_report = (
            item.contact_acquisition_report
            or (
                "CONTACT ACQUISITION REPORT\n\n"
                f"Person: {item.contact_name}\n"
                f"Title: {item.contact_title}\n"
                f"Company: {item.company_name}\n\n"
                f"Discovery method: "
                f"{item.contact_acquisition_method or source_type}\n"
                f"Primary source: {source_urls}\n\n"
                f"Email: {item.contact_email or 'Not found'}\n"
                f"Email status: {email_status}\n"
                f"Predicted pattern: "
                f"{item.contact_predicted_email_pattern or 'Not used'}\n"
                f"Pattern support: "
                f"{item.contact_pattern_support_count or 0} example(s)\n"
                f"Verification status: "
                f"{item.contact_email_verification_status or 'Unknown'}\n"
                f"Recovery method: "
                f"{item.contact_email_recovery_method or 'Not used'}\n"
                f"Domain mail status: "
                f"{item.contact_domain_mail_status or 'Not checked'}\n"
                f"Domain detail: "
                f"{item.contact_domain_mail_detail or 'No detail'}\n"
                f"Deep recovery sources: "
                f"{item.contact_deep_recovery_sources or 'None'}\n\n"
                f"Phone: "
                f"{item.contact_phone or item.contact_public_company_phone or 'Not found'}\n"
                f"Phone status: {phone_status}\n\n"
                "Important: predicted email addresses are not verified."
            )
        )
        acquisition_report = (
            acquisition_report
            + "\n\n"
            + format_email_intelligence_report(email_intelligence)
            + "\n\nZOOMINFO AVAILABILITY CHECK\n\n"
            + "Availability: "
            + item.contact_zoominfo_email_availability
            + "\nEvidence: "
            + (
                item.contact_zoominfo_email_availability_detail
                or "No definitive availability detail returned."
            )
            + "\nEnrichment attempted: "
            + (
                "YES"
                if item.contact_zoominfo_enrichment_attempted
                else "NO"
            )
            + "\nEnrichment result: "
            + (
                item.contact_zoominfo_enrichment_result
                or "No enrichment attempted."
            )
            + "\n\n"
            + format_company_intelligence(
                company_brief,
                email_resolution,
            )
            + "\n\nDiagnostic log folder: "
            + str(LOG_DIR / "email_diagnostics")
            + "\nUnknown availability policy: enrich only the highest-ranked "
            + "missing-email contact, not every finalist."
            + "\nEnrichment matching policy: use first name, last name, "
            + "company name, and job title first. Person ID is no longer "
            + "required for the initial enrichment attempt."
            + "\nMatch status: "
            + (item.contact_zoominfo_match_status or "Not recorded")
            + "\nZoomInfo enriched company ID: "
            + (
                item.contact_zoominfo_enriched_company_id
                or "Not recorded"
            )
            + "\nZoomInfo retry message: "
            + (
                item.contact_zoominfo_retry_message
                or "Not recorded"
            )
        )
        self.contact_acquisition_text.insert(
            "1.0",
            acquisition_report,
        )
        item.contact_decision_confidence = decision_confidence
        save_outreach_queue(self.review_queue)
        self._populate_delivery_item(item)

    def _retry_selected_contact_enrichment(self):
        """
        Enrich an existing Deal Desk contact without rerunning Discovery.

        Updates the current queue item and preserves duplicate protection.
        """
        queue_id = self.review_queue_id.get()
        item = self._queue_item_by_id(queue_id)
        if not item:
            messagebox.showinfo(
                APP_TITLE,
                "Select a contact in Deal Desk first.",
            )
            return

        if item.contact_email:
            if not messagebox.askyesno(
                APP_TITLE,
                (
                    f"{item.contact_name} already has "
                    f"{item.contact_email}.\n\nRetry ZoomInfo anyway?"
                ),
            ):
                return

        first_name = item.contact_first_name.strip()
        last_name = item.contact_last_name.strip()
        if not first_name and item.contact_name:
            first_name = item.contact_name.split(" ", 1)[0]
        if not last_name and " " in item.contact_name:
            last_name = item.contact_name.split(" ", 1)[1]

        if not item.contact_id and not (first_name and last_name):
            messagebox.showerror(
                APP_TITLE,
                (
                    "Compass needs a ZoomInfo person ID or the contact's "
                    "first and last name."
                ),
            )
            return

        try:
            logger = lambda msg: None
            adapter = self._zoominfo_adapter(logger)
            tools = adapter.discover_tools()
            if "enrich_contacts" not in tools:
                raise RuntimeError(
                    "ZoomInfo MCP does not expose enrich_contacts."
                )

            contact_payload = {
                "firstName": first_name,
                "lastName": last_name,
                "companyName": item.company_name,
                "jobTitle": item.contact_title,
            }

            enrich_payload = {
                "contacts": [contact_payload],
                "requiredFields": [
                    "firstName",
                    "lastName",
                    "email",
                    "phone",
                    "mobilePhone",
                    "directPhoneDoNotCall",
                    "mobilePhoneDoNotCall",
                    "jobTitle",
                    "jobFunction",
                    "managementLevel",
                    "externalUrls",
                    "contactAccuracyScore",
                    "zoominfoCompanyId",
                    "companyName",
                ],
                "userIntent": (
                    "Replay enrichment for one existing Darwill Deal Desk "
                    "contact. Return the verified business email when "
                    "available. Match by first name, last name, company "
                    "name, and job title. Do not require a person ID."
                ),
            }

            diagnostics_root = LOG_DIR / "email_diagnostics"
            diagnostics_root.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            slug = re.sub(
                r"[^a-z0-9]+",
                "_",
                f"{item.company_name}_{item.contact_name}".lower(),
            ).strip("_")
            batch_dir = diagnostics_root / f"{stamp}_{slug}_replay"
            batch_dir.mkdir(parents=True, exist_ok=True)

            (batch_dir / "enrich_tool_schema.json").write_text(
                json.dumps(
                    tools.get("enrich_contacts", {}),
                    indent=2,
                    default=str,
                ),
                encoding="utf-8",
            )
            (batch_dir / "enrich_request.json").write_text(
                json.dumps(enrich_payload, indent=2, default=str),
                encoding="utf-8",
            )

            self.status.set(
                f"Retrying ZoomInfo email for {item.contact_name}..."
            )
            self.update_idletasks()

            adapter_result = adapter.call(
                "enrich_contacts",
                enrich_payload,
                retry_count=0,
            )
            response = adapter_result.response
            if adapter_result.classification == "limit_exceeded":
                raise ZoomInfoAdapterError(
                    "ZoomInfo MCP enrichment limit exceeded. Compass will "
                    "not retry enrichment in this session."
                )
            (batch_dir / "enrich_response.json").write_text(
                json.dumps(response, indent=2, default=str),
                encoding="utf-8",
            )

            temp_contact = RankedContact(
                company_id=str(item.company_id or ""),
                company_name=item.company_name,
                contact_id=str(item.contact_id or ""),
                first_name=first_name,
                last_name=last_name,
                title=item.contact_title,
                email="",
                direct_phone="",
                mobile_phone="",
                linkedin_url="",
                crm_excluded=False,
                rank="",
                contact_score=0,
                recommendation_reason="",
                live_research_summary="",
                research_sources="",
                contact_data_status="",
                recommendation_confidence="",
                source_type="ZoomInfo",
                source_url="",
                email_status="Missing",
                phone_status="Missing",
                email_confidence=0,
                phone_confidence=0,
                decision_maker_confidence=0,
                public_company_phone="",
                predicted_email_pattern="",
            )
            result_record = best_enrichment_record(
                response,
                temp_contact,
            )
            if result_record:
                apply_enrichment_record(
                    temp_contact,
                    result_record,
                )

            email = temp_contact.email
            email_path = (
                "contact_*.data.email"
                if email
                else recursive_email_value(response)[1]
            )
            email_fields = collect_email_diagnostics(response)
            classification = classify_enrichment_failure(
                response,
                person_id=str(item.contact_id or ""),
                extracted_email=email,
            )

            lines = [
                f"Contact: {item.contact_name}",
                f"Title: {item.contact_title}",
                f"Company: {item.company_name}",
                f"Stored ZoomInfo person ID: {item.contact_id or '<none>'}",
                "Replay enrichment called: YES",
                f"Email returned: {email or 'NO'}",
                f"Email source path: {email_path or '<none>'}",
                f"Result classification: {classification}",
                f"Match status: {temp_contact.zoominfo_match_status or '<none>'}",
                (
                    "ZoomInfo company ID: "
                    f"{temp_contact.zoominfo_enriched_company_id or '<none>'}"
                ),
                (
                    "Retry message: "
                    f"{temp_contact.zoominfo_retry_message or '<none>'}"
                ),
                (
                    "Warnings: "
                    f"{temp_contact.zoominfo_enrichment_warnings or '<none>'}"
                ),
                "Email-related response fields:",
            ]
            lines.extend(
                [
                    f"{field['path']}={field['value'] or '<empty>'}"
                    for field in email_fields
                ]
                or ["<none found>"]
            )
            (batch_dir / "diagnostic_summary.txt").write_text(
                "\n".join(lines),
                encoding="utf-8",
            )

            item.contact_zoominfo_enrichment_attempted = True
            item.contact_zoominfo_enrichment_result = classification
            item.updated_at = now_iso()

            if email:
                item.contact_email = email
                item.contact_email_status = "ZoomInfo replay returned"
                item.contact_email_confidence = (
                    temp_contact.email_confidence or 98
                )
                item.contact_email_verification_status = (
                    "ZoomInfo Enriched"
                )
                item.contact_email_recovery_method = (
                    f"Deal Desk replay: {email_path}"
                )
                item.contact_data_status = "Email recovered"

            if temp_contact.mobile_phone:
                item.contact_phone = temp_contact.mobile_phone
                item.contact_phone_status = "ZoomInfo replay returned"
                item.contact_phone_confidence = (
                    temp_contact.phone_confidence or 96
                )
            elif temp_contact.direct_phone:
                item.contact_phone = temp_contact.direct_phone
                item.contact_phone_status = "ZoomInfo replay returned"
                item.contact_phone_confidence = (
                    temp_contact.phone_confidence or 94
                )

            item.contact_zoominfo_match_status = (
                temp_contact.zoominfo_match_status
            )
            item.contact_zoominfo_retry_message = (
                temp_contact.zoominfo_retry_message
            )
            item.contact_zoominfo_enriched_company_id = (
                temp_contact.zoominfo_enriched_company_id
            )
            item.contact_zoominfo_enrichment_warnings = (
                temp_contact.zoominfo_enrichment_warnings
            )

            save_outreach_queue(self.review_queue)
            self._refresh_deal_desk()
            if self.deal_tree.exists(queue_id):
                self.deal_tree.selection_set(queue_id)
                self.deal_tree.see(queue_id)
                self._load_selected_queue_item()

            if email:
                self.status.set(
                    f"Recovered {email} for {item.contact_name}."
                )
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        f"ZoomInfo returned:\n\n{email}\n\n"
                        f"Diagnostics:\n{batch_dir}"
                    ),
                )
            else:
                self.status.set(
                    f"ZoomInfo returned no email for {item.contact_name}."
                )
                messagebox.showwarning(
                    APP_TITLE,
                    (
                        "ZoomInfo enrichment ran but returned no email.\n\n"
                        f"Result: {classification}\n\n"
                        f"Diagnostics:\n{batch_dir}"
                    ),
                )
        except Exception as exc:
            self.status.set("Replay enrichment failed.")
            try:
                if "batch_dir" in locals():
                    (batch_dir / "enrichment_error.txt").write_text(
                        str(exc),
                        encoding="utf-8",
                    )
            except Exception:
                pass
            messagebox.showerror(
                APP_TITLE,
                f"Replay enrichment failed:\n\n{exc}",
            )


    def _mcp_explorer_connection(self):
        logger = lambda message: None
        oauth = OAuthManager(
            self.client_id.get(),
            self.client_secret.get(),
            logger,
        )
        token = oauth.get_access_token()
        return ZoomInfoMCP(token, logger)

    def _zoominfo_adapter(self, logger=None):
        """Return a centralized ZoomInfo MCP adapter."""
        adapter_logger = logger or (lambda message: None)
        oauth = OAuthManager(
            self.client_id.get(),
            self.client_secret.get(),
            adapter_logger,
        )
        token = oauth.get_access_token()
        return ZoomInfoAdapter(
            ZoomInfoMCP(token, adapter_logger),
            log_root=LOG_DIR,
            logger=adapter_logger,
        )

    @staticmethod
    def _mcp_schema_for_tool(tool_value):
        if isinstance(tool_value, dict):
            return tool_value
        result = {}
        for attribute in (
            "name",
            "description",
            "inputSchema",
            "input_schema",
            "schema",
        ):
            value = getattr(tool_value, attribute, None)
            if value is not None:
                result[attribute] = value
        return result or {"value": str(tool_value)}

    def _open_mcp_explorer(self):
        """Open the standalone ZoomInfo MCP developer window."""
        open_mcp_explorer_window(
            parent=self,
            app_title=APP_TITLE,
            log_dir=LOG_DIR,
            connection_factory=self._mcp_explorer_connection,
        )



    def _selected_deal_desk_item(self):
        queue_id = self.review_queue_id.get()
        item = self._queue_item_by_id(queue_id)
        if not item:
            messagebox.showinfo(
                APP_TITLE,
                "Select a company/contact in the Deal Desk first.",
            )
            return None
        return item

    def _selected_company_items(self):
        selected = self._selected_deal_desk_item()
        if not selected:
            return []
        company_id = str(selected.company_id or "").strip()
        website = normalize_domain(selected.company_website)
        company_name = normalize_company_name(selected.company_name)

        matches = []
        for item in self.review_queue:
            same_company = False
            if company_id and str(item.company_id or "").strip() == company_id:
                same_company = True
            elif website and normalize_domain(item.company_website) == website:
                same_company = True
            elif (
                company_name
                and normalize_company_name(item.company_name) == company_name
            ):
                same_company = True
            if same_company:
                matches.append(item)
        return matches

    def _show_email_review_tab(self):
        item = self._selected_deal_desk_item()
        if not item:
            return
        try:
            self.review_notebook.select(0)
        except Exception:
            pass

    def _approve_selected_company(self):
        selected = self._selected_deal_desk_item()
        if not selected:
            return
        items = self._selected_company_items()
        if not items:
            return

        self._save_queue_edits(show_message=False)
        analyses = [
            self._update_item_deliverability(item)
            for item in items
        ]
        below_threshold = [
            item
            for item, analysis in zip(items, analyses)
            if not analysis["send_ready"]
        ]

        message = (
            f"Approve {selected.company_name} and all "
            f"{len(items)} associated contact record(s)?\n\n"
            "Approval updates the Master Database and makes the company "
            "eligible for HubSpot sync. It does not send email or enroll "
            "any contact."
        )
        if below_threshold:
            message += (
                f"\n\n{len(below_threshold)} contact(s) are below the "
                f"{self.minimum_inbox_score.get()} inbox-readiness threshold. "
                "They will remain reviewable after approval."
            )

        if not messagebox.askyesno(
            APP_TITLE,
            message,
            icon="question",
        ):
            return

        now = now_iso()
        for item in items:
            item.status = "Approved"
            item.updated_at = now

        HistoryDB(DB_PATH).update_master_lifecycle(
            selected.company_id,
            selected.company_website,
            selected.company_name,
            "Approved",
            notes=selected.reviewer_notes,
        )

        try:
            key = self.intelligence_store.upsert_queue_item(selected)
            self.intelligence_store.append_timeline(
                key,
                "company_approved",
                "Company and associated contacts approved",
                "deal_desk",
                {
                    "company": selected.company_name,
                    "contact_count": len(items),
                },
            )
        except Exception:
            pass

        save_outreach_queue(self.review_queue)
        self._refresh_master_database()
        if (
            self.master_csv_auto_sync.get()
            and self.master_csv_path.get().strip()
        ):
            self._sync_master_csv_now(show_message=False)
        self._refresh_deal_desk()

        if self.deal_tree.exists(selected.queue_id):
            self.deal_tree.selection_set(selected.queue_id)
            self.deal_tree.see(selected.queue_id)
            self._load_selected_queue_item()

        messagebox.showinfo(
            APP_TITLE,
            (
                f"{selected.company_name} approved.\n\n"
                f"{len(items)} contact record(s) are now approved."
            ),
        )

    def _mark_selected_company_for_verification(self):
        selected = self._selected_deal_desk_item()
        if not selected:
            return
        items = self._selected_company_items()
        if not items:
            return

        now = now_iso()
        for item in items:
            item.status = "Needs Verification"
            item.updated_at = now

        HistoryDB(DB_PATH).update_master_lifecycle(
            selected.company_id,
            selected.company_website,
            selected.company_name,
            "Needs Verification",
            notes=selected.reviewer_notes,
        )
        save_outreach_queue(self.review_queue)
        self._refresh_master_database()
        self._refresh_deal_desk()

        if self.deal_tree.exists(selected.queue_id):
            self.deal_tree.selection_set(selected.queue_id)
            self.deal_tree.see(selected.queue_id)
            self._load_selected_queue_item()

        messagebox.showinfo(
            APP_TITLE,
            (
                f"{selected.company_name} moved to Needs Verification.\n\n"
                f"{len(items)} contact record(s) were updated."
            ),
        )

    def _reject_selected_company(self):
        selected = self._selected_deal_desk_item()
        if not selected:
            return
        items = self._selected_company_items()
        if not items:
            return

        if not messagebox.askyesno(
            APP_TITLE,
            (
                f"Reject {selected.company_name} and all "
                f"{len(items)} associated contact record(s)?\n\n"
                "The company will remain in the Master Database as rejected "
                "so Discovery will not surface it again."
            ),
            icon="warning",
        ):
            return

        now = now_iso()
        for item in items:
            item.status = "Rejected"
            item.updated_at = now

        HistoryDB(DB_PATH).update_master_lifecycle(
            selected.company_id,
            selected.company_website,
            selected.company_name,
            "Rejected",
            notes=selected.reviewer_notes,
        )

        try:
            key = self.intelligence_store.upsert_queue_item(selected)
            self.intelligence_store.append_timeline(
                key,
                "company_rejected",
                "Company and associated contacts rejected",
                "deal_desk",
                {
                    "company": selected.company_name,
                    "contact_count": len(items),
                },
            )
        except Exception:
            pass

        save_outreach_queue(self.review_queue)
        self._refresh_master_database()
        if (
            self.master_csv_auto_sync.get()
            and self.master_csv_path.get().strip()
        ):
            self._sync_master_csv_now(show_message=False)
        self._refresh_deal_desk()

        messagebox.showinfo(
            APP_TITLE,
            (
                f"{selected.company_name} rejected.\n\n"
                f"{len(items)} contact record(s) were updated."
            ),
        )

    def _sync_selected_company_to_hubspot(self):
        selected = self._selected_deal_desk_item()
        if not selected:
            return
        items = self._selected_company_items()
        if not items:
            return

        approved = [
            item
            for item in items
            if item.status == "Approved"
        ]
        if not approved:
            messagebox.showwarning(
                APP_TITLE,
                (
                    "Approve the company before syncing it to HubSpot."
                ),
            )
            return

        if not messagebox.askyesno(
            APP_TITLE,
            (
                f"Sync {selected.company_name} and "
                f"{len(approved)} approved contact record(s) to HubSpot?\n\n"
                "This creates or reuses company/contact records and adds "
                "Compass research notes. It does not enroll contacts or "
                "send email."
            ),
        ):
            return

        try:
            client = self._hubspot_client()
            failures = []
            for item in approved:
                if item.hubspot_status == "Synced":
                    continue
                try:
                    client.sync_item(item)
                    HistoryDB(DB_PATH).update_master_lifecycle(
                        item.company_id,
                        item.company_website,
                        item.company_name,
                        "Synced to HubSpot",
                        hubspot_company_id=item.hubspot_company_id,
                    )
                    try:
                        key = self.intelligence_store.upsert_queue_item(item)
                        self.intelligence_store.append_timeline(
                            key,
                            "hubspot_synced",
                            "Company/contact synchronized to HubSpot",
                            "hubspot",
                            {
                                "hubspot_company_id": item.hubspot_company_id,
                                "hubspot_contact_id": item.hubspot_contact_id,
                            },
                        )
                    except Exception:
                        pass
                except Exception as exc:
                    item.hubspot_status = "Sync Failed"
                    failures.append(
                        f"{item.contact_name}: {exc}"
                    )

            save_outreach_queue(self.review_queue)
            self._refresh_deal_desk()
            self._refresh_master_database()
            if hasattr(self, "_refresh_hubspot_workspace"):
                self._refresh_hubspot_workspace()

            if failures:
                messagebox.showwarning(
                    APP_TITLE,
                    "Company sync completed with failures:\n\n"
                    + "\n".join(failures[:10]),
                )
            else:
                messagebox.showinfo(
                    APP_TITLE,
                    (
                        f"{selected.company_name} synchronized to HubSpot.\n\n"
                        "No contacts were enrolled and no email was sent."
                    ),
                )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE,
                f"HubSpot synchronization failed:\n\n{exc}",
            )


    def _set_initial_review_vertical_split(self):
        """Give the tabs most of the right-side vertical workspace."""
        pane = getattr(self, "review_vertical_pane", None)
        if pane is None:
            return
        try:
            height = max(1, pane.winfo_height())
            # About 31% for the contact/company summary, 69% for the tabs.
            pane.sashpos(0, max(170, int(height * 0.31)))
        except Exception:
            pass

    def _toggle_deal_desk_tabs(self):
        """Expand the tab workspace or restore the balanced split."""
        pane = getattr(self, "review_vertical_pane", None)
        if pane is None:
            return

        try:
            height = max(1, pane.winfo_height())
            self.review_tabs_expanded = not getattr(
                self,
                "review_tabs_expanded",
                False,
            )
            if self.review_tabs_expanded:
                # Preserve a compact identity strip while maximizing tabs.
                pane.sashpos(0, 92)
                if hasattr(self, "expand_tabs_button"):
                    self.expand_tabs_button.configure(
                        text="Restore Summary"
                    )
            else:
                pane.sashpos(0, max(170, int(height * 0.31)))
                if hasattr(self, "expand_tabs_button"):
                    self.expand_tabs_button.configure(
                        text="Expand Tabs"
                    )
        except Exception:
            pass


    def _save_queue_edits(self, show_message: bool = True):
        queue_id = self.review_queue_id.get()
        item = self._queue_item_by_id(queue_id)
        if not item:
            if show_message:
                messagebox.showinfo(
                    APP_TITLE, "Select an item in the review queue."
                )
            return
        item.subject_line = self.review_subject.get().strip()
        item.email_body = self.review_email_body.get(
            "1.0", "end"
        ).strip()
        item.reviewer_notes = self.review_notes.get().strip()
        item.updated_at = now_iso()
        self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)
        if show_message:
            messagebox.showinfo(APP_TITLE, "Edits saved.")

    def _set_queue_status(self, status: str):
        queue_id = self.review_queue_id.get()
        item = self._queue_item_by_id(queue_id)
        if not item:
            messagebox.showinfo(
                APP_TITLE, "Select an item in the review queue."
            )
            return
        self._save_queue_edits(show_message=False)
        analysis = self._update_item_deliverability(item)
        if status == "Approved" and not analysis["send_ready"]:
            if not messagebox.askyesno(
                APP_TITLE,
                (
                    f"This draft scores {analysis['inbox_readiness_score']}/100, "
                    f"below the enrollment threshold of "
                    f"{self.minimum_inbox_score.get()}.\n\n"
                    "Approve it for continued review anyway?\n\n"
                    "Approval does not send or enroll the contact."
                ),
                icon="warning",
            ):
                return
        item.status = status
        item.updated_at = now_iso()
        master_status = "Approved" if status == "Approved" else "Rejected"
        HistoryDB(DB_PATH).update_master_lifecycle(
            item.company_id,
            item.company_website,
            item.company_name,
            master_status,
            notes=item.reviewer_notes,
        )
        save_outreach_queue(self.review_queue)
        self._refresh_master_database()
        if (
            self.master_csv_auto_sync.get()
            and self.master_csv_path.get().strip()
        ):
            self._sync_master_csv_now(show_message=False)
        self._refresh_deal_desk()
        if self.deal_tree.exists(queue_id):
            self.deal_tree.selection_set(queue_id)
            self.deal_tree.see(queue_id)
            self._load_selected_queue_item()

    def _hubspot_client(self) -> HubSpotClient:
        token = self.hubspot_token.get().strip()
        if not token:
            raise RuntimeError(
                "Enter a HubSpot access token or service key."
            )
        keyring.set_password(SERVICE, "hubspot_token", token)
        return HubSpotClient(token)

    def _test_hubspot(self):
        try:
            self._hubspot_client().test()
            self.hubspot_connection_status.set("Connected")
            self.hubspot_workspace_message.set(
                "HubSpot connection succeeded. No records were changed."
            )
            messagebox.showinfo(
                APP_TITLE,
                "HubSpot connection succeeded. No records were changed.",
            )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE, f"HubSpot connection failed:\n\n{exc}"
            )

    def _load_hubspot_sequences(self):
        try:
            client = self._hubspot_client()
            sequences = client.list_sequences()
            self.hubspot_sequences = sequences
            labels = []
            for row in sequences:
                sequence_id = str(
                    row.get("id")
                    or row.get("sequenceId")
                    or row.get("uid")
                    or ""
                )
                name = str(
                    row.get("name")
                    or row.get("label")
                    or f"Sequence {sequence_id}"
                )
                if sequence_id:
                    labels.append(f"{name} [{sequence_id}]")
            self.sequence_combo["values"] = labels
            if hasattr(self, "hubspot_workspace_sequence_combo"):
                self.hubspot_workspace_sequence_combo["values"] = labels
            if labels:
                self.selected_sequence.set(labels[0])
            self.hubspot_connection_status.set(
                f"Connected · {len(labels)} sequence(s) available"
            )
            self.hubspot_workspace_message.set(
                f"Loaded {len(labels)} HubSpot sequence(s)."
            )
            messagebox.showinfo(
                APP_TITLE,
                f"Loaded {len(labels)} HubSpot sequences.",
            )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE, f"Could not load sequences:\n\n{exc}"
            )

    def _refresh_hubspot_workspace(self):
        if not hasattr(self, "hubspot_workspace_tree"):
            return
        self.review_queue = load_outreach_queue()
        tree = self.hubspot_workspace_tree
        tree.delete(*tree.get_children())
        counts = {"approved": 0, "ready": 0, "synced": 0, "enrolled": 0, "failed": 0}

        for item in self.review_queue:
            if item.status == "Approved":
                counts["approved"] += 1
            if item.status == "Approved" and item.hubspot_status != "Synced":
                counts["ready"] += 1
            if item.hubspot_status == "Synced":
                counts["synced"] += 1
            if item.enrollment_status == "Enrolled":
                counts["enrolled"] += 1
            if "Failed" in item.hubspot_status or "Failed" in item.enrollment_status:
                counts["failed"] += 1

            if (
                item.status != "Approved"
                and item.hubspot_status != "Synced"
                and item.enrollment_status != "Enrolled"
            ):
                continue

            tree.insert(
                "", "end", iid=item.queue_id,
                values=(
                    item.company_name,
                    f"{item.contact_name} — {item.contact_title}",
                    item.contact_email or "Missing",
                    item.status,
                    item.hubspot_status,
                    item.enrollment_status,
                ),
            )

        for key, value in counts.items():
            self.hubspot_kpi_vars[key].set(str(value))
        self.hubspot_workspace_message.set(
            f"{counts['ready']} approved record(s) ready to sync; "
            f"{counts['synced']} synced; {counts['enrolled']} enrolled."
        )

    def _sync_selected_to_hubspot(self):
        if not hasattr(self, "hubspot_workspace_tree"):
            return
        selected = self.hubspot_workspace_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a HubSpot Workspace record first.")
            return

        items = [self._queue_item_by_id(qid) for qid in selected]
        items = [item for item in items if item]
        if any(item.status != "Approved" for item in items):
            messagebox.showwarning(APP_TITLE, "Only approved records can be synchronized.")
            return

        if not messagebox.askyesno(
            APP_TITLE,
            f"Sync {len(items)} selected record(s) to HubSpot?\n\n"
            "This creates or reuses company/contact records, associates them, "
            "and adds the approved Compass research as a note.\n\n"
            "This does not enroll contacts or send email.",
        ):
            return

        try:
            client = self._hubspot_client()
            failures = []
            for item in items:
                if item.hubspot_status == "Synced":
                    continue
                try:
                    client.sync_item(item)
                    try:
                        key = self.intelligence_store.upsert_queue_item(item)
                        self.intelligence_store.append_timeline(
                            key,
                            "hubspot_synced",
                            "Company and contact synchronized to HubSpot",
                            "hubspot",
                            {
                                "hubspot_company_id": item.hubspot_company_id,
                                "hubspot_contact_id": item.hubspot_contact_id,
                            },
                        )
                    except Exception:
                        pass
                    HistoryDB(DB_PATH).update_master_lifecycle(
                        item.company_id, item.company_website, item.company_name,
                        "Synced to HubSpot",
                        hubspot_company_id=item.hubspot_company_id,
                    )
                except Exception as exc:
                    item.hubspot_status = "Sync Failed"
                    item.reviewer_notes = (
                        item.reviewer_notes + f" | HubSpot sync error: {exc}"
                    ).strip(" |")
                    failures.append(f"{item.company_name}: {exc}")

            save_outreach_queue(self.review_queue)
            self._refresh_deal_desk()
            self._refresh_master_database()
            self._refresh_hubspot_workspace()
            if failures:
                messagebox.showwarning(
                    APP_TITLE,
                    "Selected sync completed with failures:\n\n"
                    + "\n".join(failures[:10]),
                )
            else:
                messagebox.showinfo(
                    APP_TITLE,
                    "Selected records synchronized. No contacts were enrolled.",
                )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE, f"HubSpot synchronization failed:\n\n{exc}"
            )

    def _approved_queue_items(self) -> list[ReviewQueueItem]:
        return [
            item for item in self.review_queue
            if item.status == "Approved"
        ]

    def _sync_approved_to_hubspot(self):
        approved = self._approved_queue_items()
        if not approved:
            messagebox.showinfo(
                APP_TITLE,
                "There are no approved items to sync.",
            )
            return
        without_email = [
            item for item in approved if not item.contact_email
        ]
        prompt = (
            f"Sync {len(approved)} approved record(s) to HubSpot?\n\n"
            "This creates or reuses companies and contacts, associates them, "
            "and adds the approved research/email as a note.\n\n"
            "This action does NOT enroll contacts or send email."
        )
        if without_email:
            prompt += (
                f"\n\n{len(without_email)} item(s) have no email and may fail "
                "contact creation."
            )
        if not messagebox.askyesno(APP_TITLE, prompt):
            return

        try:
            client = self._hubspot_client()
            failures = []
            for item in approved:
                if item.hubspot_status == "Synced":
                    continue
                try:
                    client.sync_item(item)
                    try:
                        key = self.intelligence_store.upsert_queue_item(item)
                        self.intelligence_store.append_timeline(
                            key,
                            "hubspot_synced",
                            "Company and contact synchronized to HubSpot",
                            "hubspot",
                            {
                                "hubspot_company_id": item.hubspot_company_id,
                                "hubspot_contact_id": item.hubspot_contact_id,
                            },
                        )
                    except Exception:
                        pass
                except Exception as exc:
                    item.hubspot_status = "Sync Failed"
                    item.reviewer_notes = (
                        item.reviewer_notes + f" | HubSpot sync error: {exc}"
                    ).strip(" |")
                    failures.append(f"{item.company_name}: {exc}")
            save_outreach_queue(self.review_queue)
            for synced_item in approved:
                if synced_item.hubspot_status == "Synced":
                    HistoryDB(DB_PATH).update_master_lifecycle(
                        synced_item.company_id,
                        synced_item.company_website,
                        synced_item.company_name,
                        "Synced to HubSpot",
                        hubspot_company_id=synced_item.hubspot_company_id,
                    )
            self._refresh_deal_desk()
            self._refresh_master_database()
            self._refresh_hubspot_workspace()
            if failures:
                messagebox.showwarning(
                    APP_TITLE,
                    "Sync completed with some failures:\n\n"
                    + "\n".join(failures[:10]),
                )
            else:
                messagebox.showinfo(
                    APP_TITLE,
                    "Approved records were synchronized. No contacts were enrolled.",
                )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE, f"HubSpot synchronization failed:\n\n{exc}"
            )

    def _selected_sequence_id_name(self) -> tuple[str, str]:
        selected = self.selected_sequence.get().strip()
        match = re.search(r"\[([^\]]+)\]\s*$", selected)
        if not match:
            return "", ""
        return match.group(1), selected[:match.start()].strip()

    def _confirm_and_enroll(self):
        sequence_id, sequence_name = self._selected_sequence_id_name()
        sender_email = self.hubspot_sender_email.get().strip()
        if not sequence_id:
            messagebox.showinfo(
                APP_TITLE,
                "Load and select a HubSpot sequence first.",
            )
            return
        if not sender_email:
            messagebox.showinfo(
                APP_TITLE,
                "Enter the HubSpot sequence sender email.",
            )
            return

        threshold = int(self.minimum_inbox_score.get())
        approved_synced = [
            item for item in self.review_queue
            if item.status == "Approved"
            and item.hubspot_status == "Synced"
            and item.hubspot_contact_id
            and item.enrollment_status != "Enrolled"
        ]
        for item in approved_synced:
            self._update_item_deliverability(item)
        save_outreach_queue(self.review_queue)

        unverified_email_items = [
            item for item in approved_synced
            if (
                not item.contact_email
                or item.contact_email_verification_status
                in {"Needs Verification", "Missing", ""}
                or item.contact_email_status == "Predicted — not verified"
            )
        ]
        if unverified_email_items and self.block_unverified_sequence_emails.get():
            names = "\n".join(
                f"• {item.contact_name} at {item.company_name}: "
                f"{item.contact_email_status or 'No email'} / "
                f"{item.contact_email_verification_status or 'unverified'}"
                for item in unverified_email_items[:12]
            )
            messagebox.showwarning(
                APP_TITLE,
                (
                    f"{len(unverified_email_items)} approved contact(s) were "
                    "blocked because their email is missing, predicted, or "
                    "not verified:\n\n"
                    f"{names}\n\n"
                    "Verify the address outside the app or replace it with a "
                    "ZoomInfo-returned or publicly listed address before enrollment."
                ),
            )
        email_eligible = [
            item for item in approved_synced
            if (
                not self.block_unverified_sequence_emails.get()
                or item not in unverified_email_items
            )
        ]
        low_score = [
            item for item in email_eligible
            if item.inbox_readiness_score < threshold
        ]
        eligible = [
            item for item in email_eligible
            if item.inbox_readiness_score >= threshold
        ]

        if low_score and not self.allow_low_score_override.get():
            names = "\n".join(
                f"• {item.contact_name} at {item.company_name}: "
                f"{item.inbox_readiness_score}/100"
                for item in low_score[:10]
            )
            messagebox.showwarning(
                APP_TITLE,
                (
                    f"{len(low_score)} approved contact(s) are below the "
                    f"{threshold}/100 enrollment threshold and were blocked:\n\n"
                    f"{names}\n\n"
                    "Improve the drafts in Deliverability Center, or explicitly "
                    "enable the low-score override."
                ),
            )

        if low_score and self.allow_low_score_override.get():
            if not messagebox.askyesno(
                APP_TITLE,
                (
                    f"LOW-SCORE OVERRIDE\n\n"
                    f"{len(low_score)} contact(s) are below {threshold}/100. "
                    "Include them in the final enrollment preview anyway?"
                ),
                icon="warning",
            ):
                low_score = []
            eligible.extend(low_score)
        if not eligible:
            messagebox.showinfo(
                APP_TITLE,
                "No approved, synced contacts are waiting for enrollment.",
            )
            return

        preview = "\n".join(
            f"• {item.contact_name} at {item.company_name}\n"
            f"  Inbox readiness: {item.inbox_readiness_score}/100\n"
            f"  Subject: {item.subject_line}"
            for item in eligible[:12]
        )
        if len(eligible) > 12:
            preview += f"\n• …and {len(eligible) - 12} more"

        warning = (
            f"FINAL OUTBOUND APPROVAL\n\n"
            f"Sequence: {sequence_name}\n"
            f"Sender: {sender_email}\n"
            f"Contacts: {len(eligible)}\n\n"
            f"{preview}\n\n"
            "Enrolling a contact can start the HubSpot sequence and its outbound "
            "steps according to the sequence schedule. Confirm only after you "
            "have reviewed every approved subject and email."
        )
        if not messagebox.askyesno(
            APP_TITLE,
            warning,
            icon="warning",
        ):
            return

        try:
            client = self._hubspot_client()
            failures = []
            for item in eligible:
                try:
                    client.enroll_contact(
                        item.hubspot_contact_id,
                        sequence_id,
                        sender_email,
                    )
                    item.sequence_id = sequence_id
                    item.sequence_name = sequence_name
                    item.enrollment_status = "Enrolled"
                    item.updated_at = now_iso()
                except Exception as exc:
                    item.enrollment_status = "Enrollment Failed"
                    item.reviewer_notes = (
                        item.reviewer_notes
                        + f" | Sequence enrollment error: {exc}"
                    ).strip(" |")
                    failures.append(f"{item.contact_name}: {exc}")
            save_outreach_queue(self.review_queue)
            for enrolled_item in eligible:
                if enrolled_item.enrollment_status == "Enrolled":
                    HistoryDB(DB_PATH).update_master_lifecycle(
                        enrolled_item.company_id,
                        enrolled_item.company_website,
                        enrolled_item.company_name,
                        "In Sequence",
                        hubspot_company_id=enrolled_item.hubspot_company_id,
                        sequence_status=enrolled_item.sequence_name,
                    )
            self._refresh_deal_desk()
            self._refresh_master_database()
            self._refresh_hubspot_workspace()
            if failures:
                messagebox.showwarning(
                    APP_TITLE,
                    "Enrollment completed with failures:\n\n"
                    + "\n".join(failures[:10]),
                )
            else:
                messagebox.showinfo(
                    APP_TITLE,
                    f"{len(eligible)} contact(s) enrolled in {sequence_name}.",
                )
        except Exception as exc:
            messagebox.showerror(
                APP_TITLE, f"Sequence enrollment failed:\n\n{exc}"
            )

    def _export_approved_hubspot_csv(self):
        approved = self._approved_queue_items()
        if not approved:
            messagebox.showinfo(
                APP_TITLE,
                "There are no approved records to export.",
            )
            return
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        path = HUBSPOT_EXPORT_DIR / f"HubSpot_Approved_{stamp}.csv"
        columns = [
            "Company name", "Company domain name", "Company city",
            "Company state/region", "Annual revenue", "Number of employees",
            "First name", "Last name", "Email", "Phone number", "Job title",
            "Darwill strategy", "Outreach confidence",
            "Inbox readiness", "Spam risk", "Subject quality",
            "Personalization", "Human tone", "Approved subject",
            "Approved email", "Deliverability issues",
            "Deliverability recommendations", "Reviewer notes",
        ]
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns)
            writer.writeheader()
            for item in approved:
                writer.writerow({
                    "Company name": item.company_name,
                    "Company domain name": normalize_domain(item.company_website),
                    "Company city": item.company_city,
                    "Company state/region": item.company_state,
                    "Annual revenue": item.company_revenue or "",
                    "Number of employees": item.company_employees or "",
                    "First name": item.contact_first_name,
                    "Last name": item.contact_last_name,
                    "Email": item.contact_email,
                    "Phone number": item.contact_phone,
                    "Job title": item.contact_title,
                    "Darwill strategy": item.recommended_strategy,
                    "Outreach confidence": item.outreach_confidence,
                    "Inbox readiness": item.inbox_readiness_score,
                    "Spam risk": item.spam_risk,
                    "Subject quality": item.subject_quality_score,
                    "Personalization": item.personalization_score,
                    "Human tone": item.human_tone_score,
                    "Approved subject": item.subject_line,
                    "Approved email": item.email_body,
                    "Deliverability issues": item.deliverability_issues,
                    "Deliverability recommendations": item.deliverability_recommendations,
                    "Reviewer notes": item.reviewer_notes,
                })
        os.startfile(str(HUBSPOT_EXPORT_DIR))
        messagebox.showinfo(
            APP_TITLE,
            f"Approved HubSpot CSV created:\n\n{path}",
        )

    def _populate_darwill_knowledge(self):
        if not hasattr(self, "knowledge_company_summary"):
            return
        self.knowledge_company_summary.delete("1.0", "end")
        self.knowledge_company_summary.insert(
            "1.0", self.darwill_knowledge.get("company_summary", "")
        )
        self.knowledge_meeting_cta.delete("1.0", "end")
        self.knowledge_meeting_cta.insert(
            "1.0", self.darwill_knowledge.get("meeting_cta", "")
        )
        offerings = self.darwill_knowledge.get("offerings", {})
        for strategy, widgets in self.knowledge_offering_widgets.items():
            value_box, proof_box = widgets
            offering = offerings.get(strategy, {})
            value_box.delete("1.0", "end")
            value_box.insert("1.0", offering.get("value_proposition", ""))
            proof_box.delete("1.0", "end")
            proof_box.insert("1.0", offering.get("proof_points", ""))

    def _save_darwill_knowledge(self):
        if not hasattr(self, "knowledge_company_summary"):
            return
        knowledge = {
            "company_summary": self.knowledge_company_summary.get(
                "1.0", "end"
            ).strip(),
            "meeting_cta": self.knowledge_meeting_cta.get(
                "1.0", "end"
            ).strip(),
            "offerings": {},
        }
        for strategy, widgets in self.knowledge_offering_widgets.items():
            value_box, proof_box = widgets
            knowledge["offerings"][strategy] = {
                "value_proposition": value_box.get("1.0", "end").strip(),
                "proof_points": proof_box.get("1.0", "end").strip(),
            }
        self.darwill_knowledge = knowledge
        DARWILL_KNOWLEDGE_FILE.write_text(
            json.dumps(knowledge, indent=2), encoding="utf-8"
        )
        messagebox.showinfo(APP_TITLE, "Darwill knowledge saved.")

    def _reset_darwill_knowledge(self):
        if not messagebox.askyesno(
            APP_TITLE,
            "Reset Darwill messaging to the built-in defaults?",
        ):
            return
        self.darwill_knowledge = json.loads(
            json.dumps(DEFAULT_DARWILL_KNOWLEDGE)
        )
        DARWILL_KNOWLEDGE_FILE.write_text(
            json.dumps(self.darwill_knowledge, indent=2),
            encoding="utf-8",
        )
        self._populate_darwill_knowledge()

    def _show_splash(self):
        splash = tk.Toplevel(self)
        splash.overrideredirect(True)
        splash.configure(bg=NAVY)
        width, height = 520, 300
        x = self.winfo_screenwidth() // 2 - width // 2
        y = self.winfo_screenheight() // 2 - height // 2
        splash.geometry(f"{width}x{height}+{x}+{y}")
        splash.attributes("-topmost", True)

        canvas = tk.Canvas(
            splash,
            width=520,
            height=300,
            bg=NAVY,
            highlightthickness=0,
        )
        canvas.pack(fill="both", expand=True)
        canvas.create_oval(
            216, 36, 304, 124,
            fill=ACCENT,
            outline="#8EC7F4",
            width=3,
        )
        canvas.create_text(
            260, 82,
            text="D",
            fill=WHITE,
            font=("Segoe UI Semibold", 38),
        )
        canvas.create_text(
            260, 160,
            text="DARWILL Compass",
            fill=WHITE,
            font=("Segoe UI Semibold", 22),
        )
        canvas.create_text(
            260, 192,
            text="Prospect Intelligence Platform",
            fill="#CFE3F6",
            font=("Segoe UI", 10),
        )
        canvas.create_text(
            260, 238,
            text=f"Developed by {DEVELOPER_NAME}",
            fill="#93B8D6",
            font=("Segoe UI", 8),
        )
        canvas.create_text(
            260, 266,
            text=f"Version {PRODUCT_VERSION}",
            fill="#6F9ABB",
            font=("Segoe UI Semibold", 8),
        )
        splash.after(1250, splash.destroy)

    def _toggle_sidebar(self):
        """Collapse or expand the navigation without losing page state."""
        if not hasattr(self, "sidebar_frame"):
            return

        self.sidebar_collapsed = not self.sidebar_collapsed
        width = (
            self.sidebar_collapsed_width
            if self.sidebar_collapsed
            else self.sidebar_expanded_width
        )
        self.sidebar_frame.configure(width=width)

        for label in self._sidebar_section_labels:
            if self.sidebar_collapsed:
                label.pack_forget()
            elif not label.winfo_manager():
                label.pack(fill="x", pady=(10, 4))

        for button in self._sidebar_buttons:
            try:
                button.configure(
                    text=(
                        button._compass_short_text
                        if self.sidebar_collapsed
                        else button._compass_full_text
                    ),
                    anchor=(
                        "center"
                        if self.sidebar_collapsed
                        else "w"
                    ),
                    padx=(
                        4
                        if self.sidebar_collapsed
                        else 22
                    ),
                )
                button.pack_configure(
                    padx=6 if self.sidebar_collapsed else 10
                )
            except (tk.TclError, AttributeError):
                continue

        self.workspace_status_var.set(
            "Navigation collapsed"
            if self.sidebar_collapsed
            else "Navigation expanded"
        )

    def _run_global_workspace_search(self):
        """Search currently loaded Deal Desk rows by company or contact."""
        query = self.global_search_var.get().strip()
        if not query:
            self.workspace_status_var.set(
                "Enter a company, contact, state, trade, or keyword"
            )
            return

        lowered = query.lower()

        try:
            if hasattr(self, "_select_workspace_page"):
                self._select_workspace_page(2)
        except Exception:
            pass

        if hasattr(self, "deal_tree"):
            for item_id in self.deal_tree.get_children():
                values = self.deal_tree.item(item_id, "values")
                haystack = " ".join(str(value) for value in values).lower()
                if lowered in haystack:
                    self.deal_tree.selection_set(item_id)
                    self.deal_tree.focus(item_id)
                    self.deal_tree.see(item_id)
                    try:
                        self._load_selected_queue_item()
                    except Exception:
                        pass
                    self.workspace_status_var.set(
                        f"Found: {query}"
                    )
                    return

        self.workspace_status_var.set(
            f"No visible Deal Desk match: {query}"
        )

    def _show_about(self):
        dialog = tk.Toplevel(self)
        dialog.title("About Darwill Compass")
        dialog.geometry("520x430")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg=SURFACE)

        top = tk.Frame(dialog, bg=NAVY, height=130)
        top.pack(fill="x")
        top.pack_propagate(False)

        logo = tk.Canvas(
            top, width=64, height=64, bg=NAVY, highlightthickness=0
        )
        logo.pack(side="left", padx=(24, 14), pady=28)
        logo.create_oval(3, 3, 61, 61, fill=ACCENT, outline="#8EC7F4", width=2)
        logo.create_text(
            32, 33, text="D", fill=WHITE,
            font=("Segoe UI Semibold", 28),
        )

        title = tk.Frame(top, bg=NAVY)
        title.pack(side="left", pady=26)
        tk.Label(
            title,
            text="DARWILL Compass",
            bg=NAVY,
            fg=WHITE,
            font=("Segoe UI Semibold", 18),
        ).pack(anchor="w")
        tk.Label(
            title,
            text=f"Version {PRODUCT_VERSION}",
            bg=NAVY,
            fg="#CFE3F6",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(3, 0))

        content = tk.Frame(dialog, bg=SURFACE)
        content.pack(fill="both", expand=True, padx=28, pady=22)

        description = (
            "A proprietary prospect intelligence platform designed to identify "
            "qualified home-service companies, research decision-makers, evaluate "
            "CRM relationships, and produce evidence-based outreach strategies."
        )
        tk.Label(
            content,
            text=description,
            bg=SURFACE,
            fg=TEXT,
            wraplength=450,
            justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        tk.Frame(content, bg=BORDER, height=1).pack(fill="x", pady=18)

        details = [
            ("Product", "Darwill Compass"),
            ("Developer", DEVELOPER_NAME),
            ("Platform", "ZoomInfo MCP + Tavily Intelligence"),
            ("Release", f"Version {PRODUCT_VERSION}"),
            ("Use", "Internal business development and prospect research"),
        ]
        for label, value in details:
            row = tk.Frame(content, bg=SURFACE)
            row.pack(fill="x", pady=4)
            tk.Label(
                row, text=label, width=12, anchor="w",
                bg=SURFACE, fg=MUTED,
                font=("Segoe UI Semibold", 9),
            ).pack(side="left")
            tk.Label(
                row, text=value, anchor="w",
                bg=SURFACE, fg=NAVY,
                font=("Segoe UI", 9),
            ).pack(side="left")

        tk.Button(
            content,
            text="Close",
            command=dialog.destroy,
            bg=ACCENT,
            fg=WHITE,
            activebackground=ACCENT_HOVER,
            activeforeground=WHITE,
            relief="flat",
            bd=0,
            padx=24,
            pady=8,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="e", pady=(24, 0))

    def _update_advanced_visibility(self):
        advanced_panel = getattr(
            self,
            "advanced_profile_panel",
            None,
        )
        if advanced_panel:
            if self.show_advanced.get():
                advanced_panel.pack(fill="x", pady=(12, 0))
            else:
                advanced_panel.pack_forget()

        widgets = [
            getattr(self, "custom_keywords_label", None),
            getattr(self, "custom_keywords_entry", None),
            getattr(self, "custom_naics_label", None),
            getattr(self, "custom_naics_entry", None),
        ]
        for widget in widgets:
            if not widget:
                continue
            if self.show_advanced.get():
                widget.pack()
            else:
                widget.pack_forget()

    def _toggle_connection_manager(self):
        panel = getattr(self, "credentials_panel", None)
        if not panel:
            return
        expanded = not self.connections_expanded.get()
        self.connections_expanded.set(expanded)
        if expanded:
            panel.pack(fill="x", pady=(12, 0))
        else:
            panel.pack_forget()

    def _normalize_profile_number(self, variable):
        parsed = parse_number(variable.get())
        if parsed is not None:
            variable.set(f"{parsed:,}")

    def _selected_state_codes(self) -> list[str]:
        return [
            value.strip().upper()
            for value in self.states.get().split(",")
            if value.strip()
        ]

    def _render_state_chips(self):
        frame = getattr(self, "state_chip_frame", None)
        if not frame:
            return
        for widget in frame.winfo_children():
            widget.destroy()

        state_codes = [
            "TX", "AZ", "CO", "OK", "NM",
            "NV", "UT", "ID", "WY", "MT",
        ]
        selected = set(self._selected_state_codes())
        self.state_chip_buttons = {}

        for index, code in enumerate(state_codes):
            is_selected = code in selected
            button = tk.Button(
                frame,
                text=("✓  " if is_selected else "") + code,
                command=lambda state=code: self._toggle_state_chip(state),
                bg=ACCENT if is_selected else "#EEF4F9",
                fg=WHITE if is_selected else NAVY,
                activebackground=ACCENT_HOVER if is_selected else BLUE_LIGHT,
                activeforeground=WHITE if is_selected else NAVY,
                relief="flat",
                bd=0,
                padx=14,
                pady=8,
                cursor="hand2",
                font=("Segoe UI Semibold", 9),
            )
            button.grid(
                row=index // 10,
                column=index % 10,
                sticky="ew",
                padx=(0, 7),
                pady=3,
            )
            self.state_chip_buttons[code] = button

        for column in range(10):
            frame.columnconfigure(column, weight=1)

    def _toggle_state_chip(self, code: str):
        selected = self._selected_state_codes()
        if code in selected:
            selected = [
                item for item in selected
                if item != code
            ]
        else:
            selected.append(code)

        preferred_order = [
            "TX", "AZ", "CO", "OK", "NM",
            "NV", "UT", "ID", "WY", "MT",
        ]
        selected_set = set(selected)
        ordered = [
            item for item in preferred_order
            if item in selected_set
        ]
        ordered.extend(
            item for item in selected
            if item not in preferred_order
        )
        self.states.set(",".join(ordered))
        self._render_state_chips()

    def _sync_state_chips(self):
        if hasattr(self, "state_chip_frame"):
            self._render_state_chips()

    def _refresh_connection_status(self):
        if hasattr(self, "zoominfo_connection_status"):
            self.zoominfo_connection_status.set(
                "●  Configured"
                if self.client_id.get().strip()
                and self.client_secret.get().strip()
                else "○  Not configured"
            )
        if hasattr(self, "tavily_connection_status"):
            self.tavily_connection_status.set(
                "●  Configured"
                if self.tavily_key.get().strip()
                else "○  Not configured"
            )


    def _save_feedback(self):
        company = self.feedback_company.get().strip()
        website = self.feedback_website.get().strip()
        reason = self.feedback_reason.get().strip()
        notes = self.feedback_notes.get().strip()
        if not company and not website:
            messagebox.showerror(
                APP_TITLE,
                "Enter a company name or website.",
            )
            return
        HistoryDB(DB_PATH).save_rejection_feedback(
            "",
            company,
            website,
            reason,
            notes,
        )

        counts = self.learning_rules.setdefault("rejection_counts", {})
        counts[reason] = int(counts.get(reason, 0)) + 1

        domain = normalize_domain(website)
        if domain and reason in {
            "Commercial only", "Manufacturer", "Distributor",
            "Retailer", "Wrong trade", "Franchise",
        }:
            blocked = self.learning_rules.setdefault("blocked_domains", [])
            if domain not in blocked:
                blocked.append(domain)

        if company and reason in {"Manufacturer", "Distributor", "Retailer"}:
            terms = self.learning_rules.setdefault("blocked_company_terms", [])
            normalized = normalize_trade_phrase(company)
            if normalized and normalized not in terms:
                terms.append(normalized)

        save_learning_rules(self.learning_rules)
        self.feedback_company.set("")
        self.feedback_website.set("")
        self.feedback_notes.set("")
        self._refresh_learning()
        messagebox.showinfo(APP_TITLE, "Rejection feedback saved.")

    def _refresh_learning(self):
        if not hasattr(self, "learning_tree"):
            return
        self.learning_tree.delete(*self.learning_tree.get_children())
        try:
            rows = HistoryDB(DB_PATH).rejection_feedback()
        except Exception:
            rows = []
        for _company_id, company, website, reason, notes, created_at in rows:
            self.learning_tree.insert(
                "", "end",
                values=(company, website, reason, notes, created_at),
            )

    def _render_trade_checkboxes(self):
        if not hasattr(self, "trade_checkbox_frame"):
            return
        for widget in self.trade_checkbox_frame.winfo_children():
            widget.destroy()

        self.trade_chip_buttons = {}
        for index, name in enumerate(sorted(self.trade_presets)):
            if name not in self.trade_vars:
                self.trade_vars[name] = tk.BooleanVar(value=False)

            selected = self.trade_vars[name].get()
            button = tk.Button(
                self.trade_checkbox_frame,
                text=("✓  " if selected else "") + name,
                command=lambda trade=name: self._toggle_trade_chip(trade),
                bg=ACCENT if selected else "#EEF4F9",
                fg=WHITE if selected else NAVY,
                activebackground=ACCENT_HOVER if selected else BLUE_LIGHT,
                activeforeground=WHITE if selected else NAVY,
                relief="flat",
                bd=0,
                padx=14,
                pady=8,
                cursor="hand2",
                font=("Segoe UI Semibold", 9),
            )
            button.grid(
                row=index // 4,
                column=index % 4,
                sticky="ew",
                padx=(0, 8),
                pady=3,
            )
            self.trade_chip_buttons[name] = button

        for col in range(4):
            self.trade_checkbox_frame.columnconfigure(col, weight=1)

    def _toggle_trade_chip(self, name: str):
        variable = self.trade_vars[name]
        variable.set(not variable.get())
        self._render_trade_checkboxes()
        self._refresh_trade_preview()


    def _save_trade_presets(self):
        TRADE_PRESETS_FILE.write_text(
            json.dumps(self.trade_presets, indent=2),
            encoding="utf-8",
        )

    def _open_trade_manager(self):
        dialog = tk.Toplevel(self)
        dialog.title("Manage Trade Presets")
        dialog.geometry("820x520")
        dialog.transient(self)
        dialog.grab_set()

        outer = ttk.Frame(dialog, padding=14)
        outer.pack(fill="both", expand=True)

        left = ttk.Frame(outer)
        left.pack(side="left", fill="y", padx=(0, 14))
        ttk.Label(left, text="Trade presets", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        listbox = tk.Listbox(left, width=28, height=20)
        listbox.pack(fill="y", expand=True, pady=(8, 0))

        right = ttk.LabelFrame(outer, text="Trade details", padding=12)
        right.pack(side="left", fill="both", expand=True)

        name_var = tk.StringVar()
        keywords_var = tk.StringVar()
        naics_var = tk.StringVar()

        ttk.Label(right, text="Display name").grid(row=0, column=0, sticky="w")
        name_entry = ttk.Entry(right, textvariable=name_var)
        name_entry.grid(row=1, column=0, sticky="ew", pady=(4, 6))
        suggest_status = tk.StringVar(value="")
        ttk.Button(
            right,
            text="Suggest Keywords and NAICS",
            style="Secondary.TButton",
            command=lambda: suggest_mapping(),
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))
        ttk.Label(
            right,
            textvariable=suggest_status,
            foreground=MUTED,
            wraplength=480,
        ).grid(row=3, column=0, sticky="w", pady=(0, 10))
        ttk.Label(right, text="Search keywords").grid(row=6, column=0, sticky="w")
        ttk.Entry(right, textvariable=keywords_var).grid(
            row=7, column=0, sticky="ew", pady=(4, 12)
        )
        ttk.Label(right, text="NAICS codes, comma-separated").grid(
            row=4, column=0, sticky="w"
        )
        ttk.Entry(right, textvariable=naics_var).grid(
            row=5, column=0, sticky="ew", pady=(4, 12)
        )
        ttk.Label(
            right,
            text=(
                "Example for Remodeling:\n"
                "Keywords: residential remodeling, home renovation, kitchen remodeling\n"
                "NAICS: 236118"
            ),
            foreground=MUTED,
            wraplength=480,
        ).grid(row=8, column=0, sticky="w")
        right.columnconfigure(0, weight=1)

        def suggest_mapping():
            suggestion = suggest_trade_mapping(name_var.get())
            keywords_var.set(str(suggestion.get("keywords", "")))
            naics_var.set(",".join(str(x) for x in suggestion.get("naics", [])))
            confidence = int(suggestion.get("confidence", 0) or 0)
            level = "High" if confidence >= 85 else "Medium" if confidence >= 60 else "Review"
            suggest_status.set(
                f"{level} confidence ({confidence}%): "
                f"{suggestion.get('note', '')}"
            )

        def refresh_list(select_name: str = ""):
            listbox.delete(0, "end")
            names = sorted(self.trade_presets)
            for item in names:
                listbox.insert("end", item)
            if select_name in names:
                idx = names.index(select_name)
                listbox.selection_set(idx)
                listbox.see(idx)

        def load_selected(_event=None):
            selection = listbox.curselection()
            if not selection:
                return
            name = listbox.get(selection[0])
            preset = self.trade_presets.get(name, {})
            name_var.set(name)
            keywords_var.set(str(preset.get("keywords", "")))
            naics_var.set(",".join(str(x) for x in preset.get("naics", [])))

        def clear_form():
            listbox.selection_clear(0, "end")
            name_var.set("")
            keywords_var.set("")
            naics_var.set("")

        def save_trade():
            new_name = name_var.get().strip()
            keywords = keywords_var.get().strip()
            codes = [x.strip() for x in naics_var.get().split(",") if x.strip()]
            if not new_name or not keywords:
                messagebox.showerror(
                    APP_TITLE,
                    "Trade name and search keywords are required.",
                    parent=dialog,
                )
                return

            selection = listbox.curselection()
            old_name = listbox.get(selection[0]) if selection else ""
            was_selected = (
                bool(self.trade_vars.get(old_name))
                and self.trade_vars[old_name].get()
            )
            if old_name and old_name != new_name:
                self.trade_presets.pop(old_name, None)
                self.trade_vars.pop(old_name, None)

            self.trade_presets[new_name] = {
                "keywords": keywords,
                "naics": codes,
            }
            if new_name not in self.trade_vars:
                self.trade_vars[new_name] = tk.BooleanVar(value=was_selected)

            self._save_trade_presets()
            self._render_trade_checkboxes()
            self._refresh_trade_preview()
            refresh_list(new_name)

        def delete_trade():
            selection = listbox.curselection()
            if not selection:
                return
            name = listbox.get(selection[0])
            if not messagebox.askyesno(
                APP_TITLE,
                f"Delete the '{name}' trade preset?",
                parent=dialog,
            ):
                return
            self.trade_presets.pop(name, None)
            self.trade_vars.pop(name, None)
            self._save_trade_presets()
            self._render_trade_checkboxes()
            self._refresh_trade_preview()
            clear_form()
            refresh_list()

        buttons = ttk.Frame(right)
        buttons.grid(row=9, column=0, sticky="ew", pady=(18, 0))
        ttk.Button(
            buttons, text="New", command=clear_form
        ).pack(side="left")
        ttk.Button(
            buttons, text="Save Trade", style="Blue.TButton",
            command=save_trade,
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            buttons, text="Delete", command=delete_trade
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            buttons, text="Close", command=dialog.destroy
        ).pack(side="right")

        listbox.bind("<<ListboxSelect>>", load_selected)
        refresh_list("Remodeling")

    def _selected_trade_names(self) -> list[str]:
        return [
            name for name, variable in self.trade_vars.items() if variable.get()
        ]

    def _refresh_trade_preview(self):
        selected = self._selected_trade_names()
        keyword_groups, codes = trade_configuration(
            selected,
            self.trade_presets,
            self.custom_trade_keywords.get(),
            self.naics_codes.get(),
        )
        self.trades.set(", ".join(keyword_groups))
        if hasattr(self, "trade_preview"):
            self.trade_preview.set(
                "Selected search terms: "
                + (", ".join(selected) if selected else "none")
                + "  |  NAICS sent to ZoomInfo: "
                + (codes or "none")
            )

    def _update_search_mode(self):
        if not hasattr(self, "states_entry"):
            return
        state_mode = self.search_mode.get() == "states"
        self.states_entry.configure(state="normal" if state_mode else "disabled")
        self.zip_entry.configure(state="disabled" if state_mode else "normal")
        self.radius_combo.configure(state="disabled" if state_mode else "readonly")

    def _profile_values(self) -> dict[str, Any]:
        return {
            "states": self.states.get(),
            "trades": self.trades.get(),
            "search_mode": self.search_mode.get(),
            "selected_trades": [
                name for name, variable in self.trade_vars.items() if variable.get()
            ],
            "custom_trade_keywords": self.custom_trade_keywords.get(),
            "minimum_revenue": parse_number(self.minimum_revenue.get()) or 0,
            "maximum_revenue": parse_number(self.maximum_revenue.get()) or 0,
            "target_count": self.target_count.get(),
            "candidate_limit": self.candidate_limit.get(),
            "contacts_per_company": self.contacts_per_company.get(),
            "minimum_fit_score": self.minimum_fit_score.get(),
            "naics_codes": self.naics_codes.get(),
            "employee_min": self.employee_min.get(),
            "employee_max": self.employee_max.get(),
            "territory_zip": self.territory_zip.get(),
            "territory_radius": self.territory_radius.get(),
            "exclude_crm": self.exclude_crm.get(),
            "enrich_final_contacts": self.enrich_final_contacts.get(),
            "use_zoominfo_ai_research": self.use_zoominfo_ai_research.get(),
        }

    def _load_profiles(self):
        try:
            self.profile_data = json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
            if not isinstance(self.profile_data, dict):
                self.profile_data = {}
        except Exception:
            self.profile_data = {}
        if "Default" not in self.profile_data:
            self.profile_data["Default"] = self._profile_values()
            PROFILES_FILE.write_text(
                json.dumps(self.profile_data, indent=2), encoding="utf-8"
            )
        self.profile_combo["values"] = sorted(self.profile_data)
        if self.profile_name.get() not in self.profile_data:
            self.profile_name.set("Default")

    def _apply_profile(self):
        values = self.profile_data.get(self.profile_name.get(), {})
        mapping = {
            "states": self.states, "trades": self.trades,
            "minimum_revenue": self.minimum_revenue,
            "maximum_revenue": self.maximum_revenue,
            "target_count": self.target_count,
            "candidate_limit": self.candidate_limit,
            "contacts_per_company": self.contacts_per_company,
            "minimum_fit_score": self.minimum_fit_score,
            "naics_codes": self.naics_codes,
            "employee_min": self.employee_min,
            "employee_max": self.employee_max,
            "territory_zip": self.territory_zip,
            "territory_radius": self.territory_radius,
            "exclude_crm": self.exclude_crm,
            "enrich_final_contacts": self.enrich_final_contacts,
            "smart_email_enrichment": self.smart_email_enrichment,
            "use_zoominfo_ai_research": self.use_zoominfo_ai_research,
        }
        for key, variable in mapping.items():
            if key in values:
                if key in {"minimum_revenue", "maximum_revenue"}:
                    variable.set(format_integer_with_commas(values[key]))
                else:
                    variable.set(values[key])
        self.search_mode.set(values.get("search_mode", self.search_mode.get()))
        self.custom_trade_keywords.set(
            values.get("custom_trade_keywords", self.custom_trade_keywords.get())
        )
        selected_trades = set(values.get("selected_trades", []))
        if selected_trades:
            for name, variable in self.trade_vars.items():
                variable.set(name in selected_trades)
        self._render_trade_checkboxes()
        self._refresh_trade_preview()
        self._sync_state_chips()
        self._update_search_mode()

    def _save_profile(self):
        name = self.profile_name.get().strip() or "Default"
        self.profile_data[name] = self._profile_values()
        PROFILES_FILE.write_text(
            json.dumps(self.profile_data, indent=2), encoding="utf-8"
        )
        self.profile_combo["values"] = sorted(self.profile_data)
        self.profile_name.set(name)
        messagebox.showinfo(APP_TITLE, f"Profile '{name}' saved.")

    def _save_profile_as(self):
        dialog = tk.Toplevel(self)
        dialog.title("Save Profile As")
        dialog.transient(self)
        dialog.grab_set()
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        value = tk.StringVar()
        ttk.Label(frame, text="Profile name").pack(anchor="w")
        entry = ttk.Entry(frame, textvariable=value, width=36)
        entry.pack(fill="x", pady=(6, 12))
        entry.focus_set()

        def save():
            name = value.get().strip()
            if not name:
                return
            self.profile_name.set(name)
            self._save_profile()
            dialog.destroy()

        ttk.Button(frame, text="Save", style="Blue.TButton", command=save).pack(side="right")
        ttk.Button(frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=(0, 8))

    def _delete_profile(self):
        name = self.profile_name.get()
        if name == "Default":
            messagebox.showinfo(APP_TITLE, "The Default profile cannot be deleted.")
            return
        if name and messagebox.askyesno(APP_TITLE, f"Delete profile '{name}'?"):
            self.profile_data.pop(name, None)
            PROFILES_FILE.write_text(
                json.dumps(self.profile_data, indent=2), encoding="utf-8"
            )
            self.profile_name.set("Default")
            self.profile_combo["values"] = sorted(self.profile_data)
            self._apply_profile()

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder.get() or str(OUTPUT_DIR))
        if folder:
            self.output_folder.set(folder)

    def _open_output_folder(self):
        folder = Path(self.output_folder.get() or OUTPUT_DIR)
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    def _create_shortcut(self):
        script = BASE_DIR.parent / "install_shortcuts.ps1"
        if not script.exists():
            messagebox.showerror(APP_TITLE, "Shortcut installer was not found.")
            return
        try:
            subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(script)],
                check=True,
            )
            messagebox.showinfo(
                APP_TITLE,
                "Desktop and Start Menu shortcuts were created.",
            )
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Shortcut creation failed: {exc}")

    def _check_updates(self):
        repo_root = BASE_DIR.parent
        if not (repo_root / ".git").exists():
            messagebox.showinfo(
                APP_TITLE,
                "Automatic updates require this folder to be cloned or published "
                "as a Git repository. You can still replace the folder manually.",
            )
            return
        try:
            result = subprocess.run(
                ["git", "-C", str(repo_root), "pull", "--ff-only"],
                capture_output=True, text=True, check=False,
            )
            messagebox.showinfo(
                APP_TITLE,
                (result.stdout or result.stderr or "Update check completed.").strip(),
            )
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Update check failed: {exc}")

    def _refresh_history(self):
        if not hasattr(self, "history_tree"):
            return
        self.history_tree.delete(*self.history_tree.get_children())
        try:
            rows = HistoryDB(DB_PATH).recent_runs()
        except Exception:
            rows = []
        for row in rows:
            started, profile, states, trades, qualified, contacts, reviewed, output, status = row
            self.history_tree.insert(
                "", "end",
                values=(
                    started or "", profile or "", states or "", trades or "",
                    qualified or 0, contacts or 0, reviewed or 0,
                    status or "", output or "",
                ),
            )

    def _open_selected_history(self):
        selected = self.history_tree.selection()
        if not selected:
            return
        values = self.history_tree.item(selected[0], "values")
        output = values[-1] if values else ""
        if output and Path(output).exists():
            os.startfile(output)
        else:
            messagebox.showinfo(APP_TITLE, "The selected workbook is no longer at that path.")

    def _tick_elapsed(self):
        if self.run_started_at:
            elapsed = max(0, int(time.time() - self.run_started_at))
            self.metric_vars["elapsed"].set(f"{elapsed // 60:02d}:{elapsed % 60:02d}")
        self.after(1000, self._tick_elapsed)

    def _save_checkpoint(
        self,
        trade: str,
        page: int,
        qualified: list[Prospect],
        contacts: list[RankedContact],
        seen_ids: set[str],
    ):
        payload = {
            "profile": self.profile_name.get(),
            "trade": trade,
            "page": page,
            "qualified": [asdict(item) for item in qualified],
            "contacts": [asdict(item) for item in contacts],
            "seen_ids": sorted(seen_ids),
            "saved_at": now_iso(),
        }
        CHECKPOINT_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _load_checkpoint(self) -> dict[str, Any]:
        if not self.resume_checkpoint.get() or not CHECKPOINT_FILE.exists():
            return {}
        try:
            data = json.loads(CHECKPOINT_FILE.read_text(encoding="utf-8"))
            if data.get("profile") == self.profile_name.get():
                return data
        except Exception:
            pass
        return {}

    def _reset_history(self):
        if DB_PATH.exists() and messagebox.askyesno(APP_TITLE, "Delete the local review history?"):
            DB_PATH.unlink()
            messagebox.showinfo(APP_TITLE, "Local history reset.")

    def clear_zoominfo_authorization(self):
        try:
            TOKEN_FILE.unlink(missing_ok=True)
            messagebox.showinfo(
                APP_TITLE,
                "Saved ZoomInfo authorization was cleared. The next connection will open the browser again.",
            )
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not clear authorization: {exc}")

    def connect_zoominfo(self):
        if not self.client_id.get() or not self.client_secret.get():
            messagebox.showerror(APP_TITLE, "Enter the MCP Client ID and Client Secret.")
            return
        self._save_credentials()
        self.status.set("Connecting to ZoomInfo…")
        self.progress.start(10)
        threading.Thread(target=self._connect_worker, daemon=True).start()

    def _connect_worker(self):
        try:
            logger = lambda msg: self.events.put(("log", msg))
            oauth = OAuthManager(self.client_id.get(), self.client_secret.get(), logger)
            token = oauth.get_access_token()
            mcp = ZoomInfoMCP(token, logger)
            zoominfo_adapter = ZoomInfoAdapter(
                mcp,
                log_root=LOG_DIR,
                logger=logger,
            )
            tools = zoominfo_adapter.discover_tools()
            (LOG_DIR / "mcp_tools.json").write_text(json.dumps(tools, indent=2), encoding="utf-8")
            self.mcp_tools = tools
            self.events.put(("log", f"Connected. ZoomInfo exposed {len(tools)} MCP tools."))
            self.events.put(("log", "Tools: " + ", ".join(sorted(tools))))
            self.events.put(("status", "ZoomInfo connected and ready."))
        except Exception as exc:
            self.events.put(("log", f"Connection error: {exc}\n{traceback.format_exc()}"))
            self.events.put(("status", "Connection failed—see activity log."))
        finally:
            self.events.put(("connected_done", None))

    def start(self):
        if not self.client_id.get() or not self.client_secret.get():
            messagebox.showerror(APP_TITLE, "Enter the ZoomInfo MCP credentials.")
            return
        if not self.tavily_key.get():
            messagebox.showerror(APP_TITLE, "Enter the Tavily API key.")
            return
        output_folder = Path(self.output_folder.get() or OUTPUT_DIR)
        output_folder.mkdir(parents=True, exist_ok=True)
        safe_profile = re.sub(r"[^A-Za-z0-9_-]+", "_", self.profile_name.get() or "Default")
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        output = output_folder / f"Darwill_Prospects_{safe_profile}_{stamp}.xlsx"
        self._save_settings()
        self.stop_requested = False
        self.run_started_at = time.time()
        for variable in self.metric_vars.values():
            if variable is not self.metric_vars["eta"]:
                variable.set("0")
        self.metric_vars["eta"].set("Calculating…")
        if hasattr(self, "funnel_counts"):
            for key in self.funnel_counts:
                self._update_funnel(key, 0)
        if hasattr(self, "preview_vars"):
            self._set_prospect_preview({
                "company": "Connecting to ZoomInfo MCP…",
                "phase": "STARTING",
                "location": "—",
                "size": "—",
                "score": "—",
                "residential": "—",
                "growth": "—",
                "contact": "No contact ranked yet",
                "detail": (
                    "Compass is preparing the saved profile, duplicate "
                    "protection, and live research services."
                ),
            })
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.progress.start(10)
        threading.Thread(target=self.worker, args=(Path(output),), daemon=True).start()

    def stop(self):
        self.stop_requested = True
        self.events.put(("log", "Stop requested; finishing the current research step."))

    def worker(self, output: Path):
        try:
            logger = lambda msg: self.events.put(("log", msg))
            oauth = OAuthManager(self.client_id.get(), self.client_secret.get(), logger)
            token = oauth.get_access_token()
            mcp = ZoomInfoMCP(token, logger)
            tools = mcp.run(mcp.discover())
            self.mcp_tools = tools
            (LOG_DIR / "mcp_tools.json").write_text(
                json.dumps(tools, indent=2), encoding="utf-8"
            )

            required_tools = ["search_companies", "search_contacts", "get_recommended_contacts"]
            missing = [name for name in required_tools if name not in tools]
            if missing:
                raise RuntimeError("Required ZoomInfo MCP tools missing: " + ", ".join(missing))

            states = self.states.get().strip()
            search_mode = self.search_mode.get()
            allowed_states = (
                allowed_state_codes(states) if search_mode == "states" else set()
            )
            selected_trade_names = self._selected_trade_names()
            search_specs = trade_search_specs(
                selected_trade_names,
                self.trade_presets,
                self.custom_trade_keywords.get(),
                self.naics_codes.get(),
            )
            if not search_specs:
                raise RuntimeError(
                    "Select at least one trade or enter custom trade keywords."
                )
            trades = [spec["name"] for spec in search_specs]
            minimum_dollars = parse_number(self.minimum_revenue.get()) or 0
            maximum_dollars = parse_number(self.maximum_revenue.get()) or 0
            # MCP revenue filters are in thousands.
            minimum_thousands = max(1, minimum_dollars // 1000)
            maximum_thousands = max(minimum_thousands, maximum_dollars // 1000)
            candidate_limit = max(1, int(self.candidate_limit.get()))
            target_count = max(1, int(self.target_count.get()))
            contact_limit = max(1, int(self.contacts_per_company.get()))
            minimum_fit = float(self.minimum_fit_score.get())
            employee_min = max(1, int(self.employee_min.get()))
            employee_max = max(employee_min, int(self.employee_max.get()))
            exclude_crm = self.exclude_crm.get()
            enrich_contacts_enabled = self.enrich_final_contacts.get()
            smart_email_enrichment_enabled = self.smart_email_enrichment.get()
            use_ai_research = self.use_zoominfo_ai_research.get()

            tavily = TavilyClient(api_key=self.tavily_key.get())
            history = HistoryDB(DB_PATH)
            self.current_run_id = history.start_run(
                self.profile_name.get(), states, ", ".join(trades)
            )
            checkpoint = self._load_checkpoint()
            qualified: list[Prospect] = [
                Prospect(**item) for item in checkpoint.get("qualified", [])
            ]
            ranked_contacts: list[RankedContact] = [
                RankedContact(**item) for item in checkpoint.get("contacts", [])
            ]
            outreach_drafts: list[OutreachDraft] = []
            rejected_rows: list[dict[str, Any]] = []
            qualification_decisions: list[dict[str, Any]] = []
            seen_company_ids: set[str] = set(checkpoint.get("seen_ids", []))
            resume_trade = checkpoint.get("trade", "")
            resume_page = int(checkpoint.get("page", 0) or 0)
            crm_company_exclusions = 0
            crm_contact_exclusions = 0
            history_exclusions = 0
            master_database_exclusions = 0
            hubspot_csv_company_exclusions = 0
            hubspot_csv_contact_exclusions = 0
            pages_called = 0
            candidates_reviewed = 0

            # Search each trade separately with its own focused NAICS mapping.
            per_trade_pages = max(
                2,
                (candidate_limit // max(1, len(search_specs)) + 99) // 100 + 1,
            )

            for spec_index, spec in enumerate(search_specs):
                trade = spec["name"]
                if self.stop_requested or len(qualified) >= target_count:
                    break
                if resume_trade:
                    resume_names = [item["name"] for item in search_specs]
                    if (
                        trade in resume_names
                        and resume_trade in resume_names
                        and resume_names.index(trade) < resume_names.index(resume_trade)
                    ):
                        continue

                for page in range(1, per_trade_pages + 1):
                    if trade == resume_trade and page <= resume_page:
                        continue
                    if self.stop_requested or len(qualified) >= target_count:
                        break

                    territory_zip = self.territory_zip.get().strip()

                    def apply_geography(arguments: dict[str, Any]):
                        if search_mode == "states":
                            arguments["state"] = states
                        else:
                            if not territory_zip:
                                raise RuntimeError(
                                    "Enter a ZIP code for radius search mode."
                                )
                            arguments["zipCode"] = territory_zip
                            arguments["zipCodeRadiusMiles"] = (
                                self.territory_radius.get().strip() or "50"
                            )

                    base_args = {
                        "country": "United States",
                        "locationSearchType": "HQ",
                        "revenueMin": minimum_thousands,
                        "revenueMax": maximum_thousands,
                        "employeeRangeMin": str(employee_min),
                        "employeeRangeMax": str(employee_max),
                        "page": page,
                        "pageSize": 100,
                        "sort": "-revenue",
                        "userIntent": (
                            "Internal Darwill prospecting for residential "
                            f"{trade} home-service companies."
                        )[:500],
                    }
                    apply_geography(base_args)

                    attempts: list[tuple[str, dict[str, Any]]] = []

                    # Attempt 1: concise trade label plus its own NAICS.
                    focused = dict(base_args)
                    focused["industryKeywords"] = spec["primary_keyword"]
                    if spec.get("naics"):
                        focused["naicsCodes"] = spec["naics"]
                    attempts.append(("focused keyword + trade NAICS", focused))

                    # Attempt 2: matching NAICS only. Tavily performs the
                    # residential/trade verification afterward.
                    if spec.get("naics"):
                        naics_only = dict(base_args)
                        naics_only["naicsCodes"] = spec["naics"]
                        attempts.append(("trade NAICS only", naics_only))

                    # Attempt 3: one concise keyword without business-model,
                    # description or NAICS constraints.
                    keyword_only = dict(base_args)
                    keyword_only["industryKeywords"] = spec["primary_keyword"]
                    attempts.append(("concise keyword only", keyword_only))

                    records = []
                    raw = {}
                    successful_attempt = ""
                    for attempt_name, company_args in attempts:
                        logger(
                            f"ZoomInfo structured search: trade={trade}, "
                            f"attempt={attempt_name}, page={page}, "
                            f"states={states if search_mode == 'states' else 'ZIP radius'}, "
                            f"revenue=${minimum_dollars:,}-${maximum_dollars:,}"
                        )
                        raw = mcp.run(
                            mcp.call("search_companies", company_args)
                        )
                        pages_called += 1
                        attempt_slug = re.sub(
                            r"[^a-z0-9]+", "_", attempt_name.lower()
                        ).strip("_")
                        log_path = LOG_DIR / (
                            f"company_search_"
                            f"{re.sub(r'[^a-z0-9]+','_',trade.lower()).strip('_')}_"
                            f"{attempt_slug}_{page}.json"
                        )
                        log_path.write_text(
                            json.dumps(raw, indent=2, default=str),
                            encoding="utf-8",
                        )
                        records = records_from_payload(raw)
                        logger(
                            f"ZoomInfo returned {len(records)} companies for "
                            f"{trade} using {attempt_name}, page {page}."
                        )
                        if records:
                            successful_attempt = attempt_name
                            break

                    if not records:
                        logger(
                            f"No ZoomInfo companies found for {trade} after "
                            f"{len(attempts)} structured attempts."
                        )
                        break

                    logger(
                        f"Using {successful_attempt} results for {trade}; "
                        "Tavily will enforce residential and trade fit."
                    )

                    for record in records:
                        if self.stop_requested or len(qualified) >= target_count:
                            break
                        prospect = company_from_record(record)
                        if not prospect.company_id or not prospect.company_name:
                            continue

                        learned_domain = normalize_domain(prospect.website)
                        blocked_domains = set(
                            self.learning_rules.get("blocked_domains", [])
                        )
                        blocked_terms = self.learning_rules.get(
                            "blocked_company_terms", []
                        )
                        normalized_company = normalize_trade_phrase(
                            prospect.company_name
                        )
                        if (
                            learned_domain and learned_domain in blocked_domains
                        ) or any(
                            term and term in normalized_company
                            for term in blocked_terms
                        ):
                            learning_reason = "Rejected by qualification-learning rules"
                            history.save(
                                prospect.company_id,
                                prospect.company_name,
                                "REJECT",
                                learning_reason,
                            )
                            qualification_decisions.append(
                                qualification_result(prospect, "Rejected", learning_reason, trade)
                            )
                            continue
                        if self.use_hubspot_csv_index.get():
                            hubspot_company = history.hubspot_company_match(
                                prospect.company_name,
                                prospect.website,
                            )
                            if hubspot_company:
                                hubspot_csv_company_exclusions += 1
                                reason = (
                                    "Already exists in imported HubSpot company index "
                                    f"({hubspot_company['match_type']}; "
                                    f"record {hubspot_company.get('record_id') or 'unknown'}; "
                                    f"source {hubspot_company.get('source_file') or 'CSV'})."
                                )
                                logger(
                                    f"HubSpot CSV index skipped "
                                    f"{prospect.company_name}: {reason}"
                                )
                                history.upsert_master_prospect(
                                    prospect,
                                    "Synced to HubSpot",
                                    "REJECT",
                                    reason,
                                    trade,
                                    self.profile_name.get(),
                                    self.current_run_id,
                                    0,
                                )
                                qualification_decisions.append(
                                    qualification_result(
                                        prospect, "Rejected", reason, trade
                                    )
                                )
                                continue

                        if self.use_master_dedup.get():
                            skip_master, skip_reason = history.should_skip_master(
                                prospect.company_id,
                                prospect.website,
                                prospect.company_name,
                                int(self.rejected_rereview_days.get()),
                                int(self.qualified_rereview_days.get()),
                            )
                            if skip_master:
                                master_database_exclusions += 1
                                logger(
                                    f"Master DB skipped {prospect.company_name}: "
                                    f"{skip_reason}"
                                )
                                continue

                        if prospect.company_id in seen_company_ids:
                            continue
                        seen_company_ids.add(prospect.company_id)
                        candidates_reviewed += 1
                        self.events.put(("metrics", {"reviewed": candidates_reviewed}))
                        self.events.put(("funnel", {
                            "reviewed": candidates_reviewed,
                        }))
                        self.events.put(("prospect_preview", {
                            "company": prospect.company_name,
                            "phase": "HARD FILTERS",
                            "location": ", ".join(
                                item for item in [prospect.city, prospect.state]
                                if item
                            ) or "Location unavailable",
                            "size": (
                                f"${prospect.revenue:,.0f} revenue · "
                                f"{prospect.employees:,} employees"
                                if prospect.revenue and prospect.employees
                                else (
                                    f"${prospect.revenue:,.0f} revenue"
                                    if prospect.revenue
                                    else (
                                        f"{prospect.employees:,} employees"
                                        if prospect.employees
                                        else "Company size unavailable"
                                    )
                                )
                            ),
                            "score": "Pending",
                            "residential": "Pending",
                            "growth": "Pending",
                            "detail": (
                                f"Evaluating {trade} fit, geography, duplicate "
                                "protection, and CRM history before live research."
                            ),
                        }))

                        returned_state = normalize_state(prospect.state)
                        if allowed_states and returned_state not in allowed_states:
                            rejected_rows.append({
                                "company_name": prospect.company_name,
                                "website": prospect.website,
                                "reason": (
                                    f"Returned state {prospect.state or 'missing'} is outside "
                                    f"allowed states {', '.join(sorted(allowed_states))}"
                                ),
                            })
                            reason = (
                                f"State outside target geography: "
                                f"{prospect.state or 'missing'}"
                            )
                            history.save(
                                prospect.company_id,
                                prospect.company_name,
                                "REJECT",
                                reason,
                            )
                            history.upsert_master_prospect(
                                prospect,
                                "Rejected",
                                "REJECT",
                                reason,
                                trade,
                                self.profile_name.get(),
                                self.current_run_id,
                                int(self.rejected_rereview_days.get()),
                            )
                            qualification_decisions.append(
                                qualification_result(prospect, "Rejected", reason, trade)
                            )
                            continue

                        if prospect.crm_excluded:
                            crm_company_exclusions += 1
                            crm_reason = "ZoomInfo indicates existing CRM relationship"
                            history.upsert_master_prospect(
                                prospect,
                                "Synced to HubSpot",
                                "REJECT",
                                crm_reason,
                                trade,
                                self.profile_name.get(),
                                self.current_run_id,
                                0,
                            )
                            qualification_decisions.append(
                                qualification_result(prospect, "Rejected", crm_reason, trade)
                            )
                            continue

                        prior = history.seen(prospect.company_id)
                        if self.skip_history.get() and prior:
                            history_exclusions += 1
                            continue

                        # Live public verification first; free of ZoomInfo AI credits.
                        self.events.put(("prospect_preview", {
                            "company": prospect.company_name,
                            "phase": "TAVILY RESEARCH",
                            "detail": (
                                "Verifying residential focus, business model, "
                                "service territory, growth activity, marketing "
                                "maturity, and technology signals."
                            ),
                        }))
                        research_company(tavily, prospect)
                        self.events.put(("funnel_add", {"researched": 1}))
                        self.events.put(("prospect_preview", {
                            "company": prospect.company_name,
                            "phase": "QUALIFICATION",
                            "score": f"{prospect.fit_score:.0f}",
                            "residential": (
                                prospect.residential_signals
                                or "No strong signal"
                            ),
                            "growth": (
                                prospect.growth_signals
                                or "No current signal"
                            ),
                            "detail": (
                                prospect.acceptance_reason
                                or prospect.company_summary
                                or "Live company research completed."
                            ),
                        }))
                        if prospect.fit_score < minimum_fit or prospect.exclusion_signals:
                            reason = prospect.acceptance_reason or "Failed live company qualification"
                            history.save(
                                prospect.company_id,
                                prospect.company_name,
                                "REJECT",
                                reason,
                            )
                            history.upsert_master_prospect(
                                prospect,
                                "Rejected",
                                "REJECT",
                                reason,
                                trade,
                                self.profile_name.get(),
                                self.current_run_id,
                                int(self.rejected_rereview_days.get()),
                            )
                            rejected_rows.append({
                                "company_name": prospect.company_name,
                                "website": prospect.website,
                                "reason": reason,
                            })
                            qualification_decisions.append(
                                qualification_result(prospect, "Rejected", reason, trade)
                            )
                            self.events.put(("funnel_add", {"rejected": 1}))
                            self.events.put(("prospect_preview", {
                                "company": prospect.company_name,
                                "phase": "REJECTED",
                                "score": f"{prospect.fit_score:.0f}",
                                "detail": reason,
                            }))
                            continue

                        # CRM relationship check through ZoomInfo Account Research, only
                        # after the company otherwise qualifies. This can consume AI credits.
                        if exclude_crm and use_ai_research and "account_research" in tools:
                            crm_query = (
                                "Determine whether this account already exists in our CRM/HubSpot, "
                                "has an active or historical opportunity, is under management, or has "
                                "meaningful prior engagement. Begin the response with CRM_EXISTING or "
                                "CRM_NEW, then briefly explain."
                            )
                            try:
                                crm_result = mcp.run(mcp.call("account_research", {
                                    "query": crm_query,
                                    "zoominfoCompanyId": int(prospect.company_id),
                                }))
                                crm_text = text_blob(crm_result).lower()
                                (LOG_DIR / f"account_research_{prospect.company_id}.json").write_text(
                                    json.dumps(crm_result, indent=2, default=str), encoding="utf-8"
                                )
                                if (
                                    "crm_existing" in crm_text
                                    or "already exists in our crm" in crm_text
                                    or "active opportunity" in crm_text
                                    or "under management" in crm_text
                                ):
                                    crm_company_exclusions += 1
                                    reason = (
                                        "ZoomInfo account research indicates "
                                        "existing CRM relationship"
                                    )
                                    history.save(
                                        prospect.company_id,
                                        prospect.company_name,
                                        "REJECT",
                                        reason,
                                    )
                                    history.upsert_master_prospect(
                                        prospect,
                                        "Synced to HubSpot",
                                        "REJECT",
                                        reason,
                                        trade,
                                        self.profile_name.get(),
                                        self.current_run_id,
                                        0,
                                    )
                                    qualification_decisions.append(
                                        qualification_result(prospect, "Rejected", reason, trade)
                                    )
                                    continue
                            except Exception as exc:
                                logger(f"CRM account check warning for {prospect.company_name}: {exc}")

                        qualified.append(prospect)
                        history.save(
                            prospect.company_id,
                            prospect.company_name,
                            "ACCEPT",
                            prospect.acceptance_reason,
                        )
                        master_prospect_key = history.upsert_master_prospect(
                            prospect,
                            "Qualified",
                            "ACCEPT",
                            prospect.acceptance_reason,
                            trade,
                            self.profile_name.get(),
                            self.current_run_id,
                            int(self.qualified_rereview_days.get()),
                        )
                        qualification_decisions.append(
                            qualification_result(prospect, "Accepted", prospect.acceptance_reason, trade)
                        )
                        logger(
                            f"Qualified {len(qualified)}/{target_count}: "
                            f"{prospect.company_name} ({prospect.fit_score:.0f})"
                        )
                        self.events.put(("funnel", {
                            "qualified": len(qualified),
                        }))
                        self.events.put(("prospect_preview", {
                            "company": prospect.company_name,
                            "phase": "QUALIFIED",
                            "score": f"{prospect.fit_score:.0f}",
                            "residential": (
                                prospect.residential_signals
                                or "Verified"
                            ),
                            "growth": (
                                prospect.growth_signals
                                or "No current signal"
                            ),
                            "detail": (
                                prospect.acceptance_reason
                                or "Qualified by the proven 3.6 engine."
                            ),
                        }))
                        elapsed = max(1.0, time.time() - self.run_started_at)
                        rate = len(qualified) / elapsed
                        remaining = max(0, target_count - len(qualified))
                        eta_seconds = int(remaining / rate) if rate > 0 else 0
                        self.events.put(("metrics", {
                            "qualified": len(qualified),
                            "contacts": len(ranked_contacts),
                            "reviewed": candidates_reviewed,
                            "eta": (
                                f"{eta_seconds // 60:02d}:{eta_seconds % 60:02d}"
                                if eta_seconds else "Calculating…"
                            ),
                        }))

                        contact_records: list[dict[str, Any]] = []

                        # ZoomInfo's ML recommendations are free and use the user's
                        # interaction/CRM history.
                        try:
                            recommendations = mcp.run(mcp.call("get_recommended_contacts", {
                                "useCaseType": "PROSPECTING",
                                "ziCompanyId": int(prospect.company_id),
                                "pageSize": 25,
                                "userIntent": (
                                    "Recommend marketing and growth decision-makers for Darwill's "
                                    "customer acquisition, direct mail, new-mover, and digital services."
                                ),
                            }))
                            (LOG_DIR / f"recommended_contacts_{prospect.company_id}.json").write_text(
                                json.dumps(recommendations, indent=2, default=str), encoding="utf-8"
                            )
                            contact_records.extend(records_from_payload(recommendations))
                        except Exception as exc:
                            logger(f"Recommended contacts warning for {prospect.company_name}: {exc}")

                        # Structured title search, exact company ID.
                        title_query = (
                            "Chief Marketing Officer OR Vice President Marketing OR "
                            "Director Marketing OR Vice President Growth OR Director Growth OR "
                            "Head Marketing OR Marketing Manager OR President OR "
                            "Chief Executive Officer OR Owner OR Chief Operating Officer OR General Manager"
                        )
                        contact_args = {
                            "companyId": prospect.company_id,
                            "jobTitle": title_query,
                            "excludeJobTitle": (
                                "Finance,Financial,Controller,Accounting,Human Resources,"
                                "Recruiter,Legal,Attorney,Procurement,Technician,Coordinator,Assistant"
                            ),
                            "page": 1,
                            "pageSize": 100,
                            "userIntent": (
                                "Find current marketing, growth, ownership, and executive "
                                "decision-makers for external marketing-services prospecting."
                            ),
                        }
                        try:
                            contact_search = mcp.run(mcp.call("search_contacts", contact_args))
                            (LOG_DIR / f"contact_search_{prospect.company_id}.json").write_text(
                                json.dumps(contact_search, indent=2, default=str), encoding="utf-8"
                            )
                            contact_records.extend(records_from_payload(contact_search))
                        except Exception as exc:
                            logger(f"Contact search warning for {prospect.company_name}: {exc}")

                        candidates: dict[str, RankedContact] = {}
                        recommendation_meta: dict[str, str] = {}
                        zoominfo_availability_meta: dict[
                            str, tuple[str, str]
                        ] = {}

                        for contact_record in contact_records:
                            contact = contact_from_record(contact_record, prospect)
                            if not contact.contact_id:
                                continue
                            availability, availability_detail = (
                                zoominfo_email_availability(contact_record)
                            )
                            prior = zoominfo_availability_meta.get(
                                contact.contact_id
                            )
                            if (
                                prior is None
                                or (
                                    prior[0] == "Unknown"
                                    and availability != "Unknown"
                                )
                                or availability == "Available"
                            ):
                                zoominfo_availability_meta[contact.contact_id] = (
                                    availability,
                                    availability_detail,
                                )
                            contact.zoominfo_email_availability = availability
                            contact.zoominfo_email_availability_detail = (
                                availability_detail
                            )
                            if contact.crm_excluded:
                                crm_contact_exclusions += 1
                                continue
                            score, reasons = score_contact(contact)
                            if score <= 0:
                                continue

                            # Preserve ZoomInfo recommendation metadata and scores.
                            meta_text = text_blob(contact_record.get("meta", ""))
                            zi_score = parse_number(first_value(contact_record, ["reRankingScore", "score"], ""))
                            if zi_score is not None:
                                score += min(20, max(0, float(zi_score)))
                                reasons.append("ZoomInfo recommendation score")
                            if meta_text:
                                recommendation_meta[contact.contact_id] = meta_text[:500]

                            contact.contact_score = score
                            contact.recommendation_reason = ", ".join(reasons)
                            existing = candidates.get(contact.contact_id)
                            if existing is None or contact.contact_score > existing.contact_score:
                                candidates[contact.contact_id] = contact

                        zoominfo_finalists = sorted(
                            candidates.values(),
                            key=lambda c: c.contact_score,
                            reverse=True,
                        )[:max(contact_limit * 5, 12)]

                        public_candidates: list[RankedContact] = []
                        public_company_phone = ""
                        if self.discover_public_contacts.get():
                            try:
                                (
                                    public_candidates,
                                    public_company_phone,
                                ) = discover_public_decision_makers(
                                    tavily,
                                    prospect,
                                )
                                logger(
                                    f"Public research found "
                                    f"{len(public_candidates)} additional "
                                    f"decision-maker candidate(s) for "
                                    f"{prospect.company_name}."
                                )
                            except Exception as exc:
                                logger(
                                    f"Public contact discovery warning for "
                                    f"{prospect.company_name}: {exc}"
                                )

                        if not self.predict_public_emails.get():
                            for public_contact in public_candidates:
                                if public_contact.email_status == "Predicted — not verified":
                                    public_contact.email = ""
                                    public_contact.email_status = "Missing"
                                    public_contact.email_confidence = 0
                                    public_contact.predicted_email_pattern = ""

                        finalists = merge_decision_maker_candidates(
                            zoominfo_finalists,
                            public_candidates,
                            prospect,
                        )[:max(contact_limit * 5, 12)]
                        for finalist in finalists:
                            availability_meta = zoominfo_availability_meta.get(
                                finalist.contact_id
                            )
                            if availability_meta:
                                (
                                    finalist.zoominfo_email_availability,
                                    finalist.zoominfo_email_availability_detail,
                                ) = availability_meta

                        if self.use_hubspot_csv_index.get():
                            deduped_finalists = []
                            for candidate in finalists:
                                hubspot_contact = history.hubspot_contact_match(
                                    candidate, prospect
                                )
                                if hubspot_contact:
                                    hubspot_csv_contact_exclusions += 1
                                    logger(
                                        f"HubSpot CSV contact index skipped "
                                        f"{candidate.first_name} "
                                        f"{candidate.last_name} at "
                                        f"{prospect.company_name}: "
                                        f"{hubspot_contact['match_type']} "
                                        f"(record "
                                        f"{hubspot_contact.get('record_id') or 'unknown'})."
                                    )
                                    continue
                                deduped_finalists.append(candidate)
                            finalists = deduped_finalists

                        for candidate in finalists:
                            if (
                                public_company_phone
                                and not candidate.public_company_phone
                            ):
                                candidate.public_company_phone = (
                                    public_company_phone
                                )
                                if not candidate.phone_status:
                                    candidate.phone_status = (
                                        "Public company phone"
                                    )
                                    candidate.phone_confidence = 75

                        # Smart ZoomInfo enrichment: free/public research first,
                        # then paid enrichment only when raw ZoomInfo metadata
                        # indicates that an email is available.
                        smart_enrichment_enabled = self.smart_email_enrichment.get()
                        enrichment_candidates: list[RankedContact] = []

                        if smart_enrichment_enabled:
                            confirmed_available = [
                                contact
                                for contact in finalists
                                if (
                                    not contact.email
                                    and contact.zoominfo_email_availability
                                    == "Available"
                                )
                            ]
                            unknown_availability = [
                                contact
                                for contact in finalists
                                if (
                                    not contact.email
                                    and contact.zoominfo_email_availability
                                    == "Unknown"
                                )
                            ]
                            enrichment_candidates = confirmed_available[:10]
                            if (
                                not enrichment_candidates
                                and unknown_availability
                            ):
                                unknown_availability.sort(
                                    key=lambda candidate: (
                                        candidate.outreach_order
                                        if candidate.outreach_order > 0
                                        else 999,
                                        -candidate.contact_score,
                                    )
                                )
                                fallback_contact = unknown_availability[0]
                                fallback_contact.zoominfo_email_availability_detail = (
                                    fallback_contact.zoominfo_email_availability_detail
                                    + " Targeted fallback selected because MCP "
                                    "did not expose availability; ZoomInfo Sales "
                                    "may still contain a visible email."
                                ).strip()
                                enrichment_candidates = [fallback_contact]

                            logger(
                                f"Smart ZoomInfo email check for "
                                f"{prospect.company_name}: "
                                f"{len(confirmed_available)} confirmed available; "
                                f"{len(unknown_availability)} unknown; "
                                f"{len(enrichment_candidates)} submitted."
                            )
                        elif enrich_contacts_enabled:
                            enrichment_candidates = [
                                contact
                                for contact in finalists
                                if (
                                    not contact.email
                                    and str(contact.contact_id).isdigit()
                                )
                            ][:10]

                        if (
                            "zoominfo_adapter" in locals()
                            and zoominfo_adapter.enrichment_blocked
                        ):
                            logger(
                                "ZoomInfo enrichment blocked for this run: "
                                + zoominfo_adapter.enrichment_block_reason
                            )
                            enrichment_candidates = []

                        if enrichment_candidates and "enrich_contacts" in tools:
                            enrich_payload = {
                                "contacts": [
                                    {
                                        "firstName": contact.first_name,
                                        "lastName": contact.last_name,
                                        "companyName": contact.company_name,
                                        "jobTitle": contact.title,
                                    }
                                    for contact in enrichment_candidates
                                ],
                                "requiredFields": [
                                    "firstName", "lastName", "email", "phone",
                                    "mobilePhone", "directPhoneDoNotCall",
                                    "mobilePhoneDoNotCall", "jobTitle",
                                    "jobFunction", "managementLevel",
                                    "externalUrls", "contactAccuracyScore",
                                    "zoominfoCompanyId", "companyName",
                                ],
                                "userIntent": (
                                    "Retrieve verified business emails only for "
                                    "shortlisted Darwill contacts whose ZoomInfo "
                                    "search results indicate email availability. "
                                    "Match primarily by first name, last name, "
                                    "company name, and job title."
                                ),
                            }
                            try:
                                for contact in enrichment_candidates:
                                    contact.zoominfo_enrichment_attempted = True
                                    contact.zoominfo_enrichment_result = (
                                        "Submitted after availability confirmation "
                                        f"using {contact.zoominfo_person_id_source}"
                                    )
                                diagnostics_root = (
                                    LOG_DIR / "email_diagnostics"
                                )
                                diagnostics_root.mkdir(
                                    parents=True,
                                    exist_ok=True,
                                )
                                request_stamp = datetime.now().strftime(
                                    "%Y%m%d_%H%M%S"
                                )
                                company_slug = re.sub(
                                    r"[^a-z0-9]+",
                                    "_",
                                    prospect.company_name.lower(),
                                ).strip("_")
                                batch_dir = (
                                    diagnostics_root
                                    / f"{request_stamp}_{company_slug}"
                                )
                                batch_dir.mkdir(
                                    parents=True,
                                    exist_ok=True,
                                )
                                (
                                    batch_dir / "enrich_tool_schema.json"
                                ).write_text(
                                    json.dumps(
                                        tools.get("enrich_contacts", {}),
                                        indent=2,
                                        default=str,
                                    ),
                                    encoding="utf-8",
                                )
                                (
                                    batch_dir / "enrich_request.json"
                                ).write_text(
                                    json.dumps(
                                        enrich_payload,
                                        indent=2,
                                        default=str,
                                    ),
                                    encoding="utf-8",
                                )

                                adapter_result = zoominfo_adapter.call(
                                    "enrich_contacts",
                                    enrich_payload,
                                    retry_count=0,
                                )
                                enriched = adapter_result.response
                                if (
                                    adapter_result.classification
                                    == "limit_exceeded"
                                ):
                                    logger(
                                        "ZoomInfo enrichment limit exceeded. "
                                        "Remaining enrichment calls will be "
                                        "skipped for this run."
                                    )
                                (
                                    batch_dir / "enrich_response.json"
                                ).write_text(
                                    json.dumps(
                                        enriched,
                                        indent=2,
                                        default=str,
                                    ),
                                    encoding="utf-8",
                                )
                                (
                                    LOG_DIR
                                    / f"enriched_contacts_{prospect.company_id}.json"
                                ).write_text(
                                    json.dumps(enriched, indent=2, default=str),
                                    encoding="utf-8",
                                )
                                enriched_by_id = {}
                                enrichment_records = records_from_payload(enriched)
                                for record in enrichment_records:
                                    enriched_contact = contact_from_record(
                                        record,
                                        prospect,
                                    )
                                    if enriched_contact.contact_id:
                                        enriched_by_id[
                                            enriched_contact.contact_id
                                        ] = enriched_contact

                                successful_records = (
                                    enrichment_result_records(enriched)
                                )
                                payload_email, payload_email_path = (
                                    recursive_email_value(enriched)
                                )
                                all_email_fields = collect_email_diagnostics(
                                    enriched
                                )

                                recovered_count = 0
                                diagnostic_summaries: list[str] = []
                                for contact in enrichment_candidates:
                                    result_record = best_enrichment_record(
                                        enriched,
                                        contact,
                                    )
                                    if result_record:
                                        apply_enrichment_record(
                                            contact,
                                            result_record,
                                        )
                                    else:
                                        result = enriched_by_id.get(
                                            contact.contact_id
                                        )
                                        if result:
                                            contact.email = (
                                                result.email or contact.email
                                            )
                                            contact.direct_phone = (
                                                result.direct_phone
                                                or contact.direct_phone
                                            )
                                            contact.mobile_phone = (
                                                result.mobile_phone
                                                or contact.mobile_phone
                                            )
                                            contact.linkedin_url = (
                                                result.linkedin_url
                                                or contact.linkedin_url
                                            )
                                    if (
                                        not contact.email
                                        and len(enrichment_candidates) == 1
                                        and payload_email
                                    ):
                                        contact.email = payload_email
                                        contact.email_recovery_method = (
                                            "ZoomInfo payload fallback: "
                                            + payload_email_path
                                        )

                                    failure_class = classify_enrichment_failure(
                                        enriched,
                                        person_id=str(contact.contact_id),
                                        extracted_email=contact.email,
                                    )
                                    email_field_lines = [
                                        (
                                            f"{field['path']}="
                                            f"{field['value'] or '<empty>'}"
                                        )
                                        for field in all_email_fields
                                    ]
                                    diagnostic_summaries.append(
                                        "\n".join([
                                            (
                                                f"Contact: "
                                                f"{contact.first_name} "
                                                f"{contact.last_name}"
                                            ).strip(),
                                            f"Title: {contact.title}",
                                            (
                                                "ZoomInfo person ID: "
                                                f"{contact.contact_id}"
                                            ),
                                            (
                                                "Outer MCP record ID: "
                                                f"{contact.zoominfo_outer_record_id}"
                                            ),
                                            (
                                                "Person ID source: "
                                                f"{contact.zoominfo_person_id_source}"
                                            ),
                                            (
                                                "Availability before enrichment: "
                                                f"{contact.zoominfo_email_availability}"
                                            ),
                                            (
                                                "Availability evidence: "
                                                f"{contact.zoominfo_email_availability_detail}"
                                            ),
                                            "Enrichment called: YES",
                                            (
                                                "Email returned: "
                                                f"{contact.email or 'NO'}"
                                            ),
                                            (
                                                "Result classification: "
                                                f"{failure_class}"
                                            ),
                                            "Email-related response fields:",
                                            *(
                                                email_field_lines
                                                or ["<none found>"]
                                            ),
                                        ])
                                    )

                                    if contact.email:
                                        recovered_count += 1
                                        contact.email_status = (
                                            "ZoomInfo enrichment returned"
                                        )
                                        contact.email_confidence = max(
                                            contact.email_confidence, 98
                                        )
                                        contact.email_verification_status = (
                                            "ZoomInfo Enriched"
                                        )
                                        contact.email_recovery_method = (
                                            "ZoomInfo paid enrichment after "
                                            "availability confirmation"
                                        )
                                        contact.zoominfo_enrichment_result = (
                                            "Verified email recovered"
                                        )
                                    else:
                                        contact.zoominfo_enrichment_result = (
                                            "Availability was indicated, but "
                                            "enrichment returned no email"
                                        )
                                (
                                    batch_dir / "diagnostic_summary.txt"
                                ).write_text(
                                    ("\n\n" + ("-" * 72) + "\n\n").join(
                                        diagnostic_summaries
                                    ),
                                    encoding="utf-8",
                                )
                                logger(
                                    f"Email diagnostics saved to {batch_dir}."
                                )
                                logger(
                                    f"Smart enrichment recovered "
                                    f"{recovered_count} email(s) from "
                                    f"{len(enrichment_candidates)} confirmed-"
                                    f"available contact(s)."
                                )
                            except Exception as exc:
                                try:
                                    if "batch_dir" in locals():
                                        (
                                            batch_dir / "enrichment_error.txt"
                                        ).write_text(
                                            str(exc),
                                            encoding="utf-8",
                                        )
                                except Exception:
                                    pass
                                for contact in enrichment_candidates:
                                    contact.zoominfo_enrichment_result = (
                                        f"Enrichment failed: {exc}"
                                    )
                                logger(
                                    f"Contact enrichment warning for "
                                    f"{prospect.company_name}: {exc}"
                                )
                        elif smart_enrichment_enabled:
                            logger(
                                f"No confirmed-available ZoomInfo emails for "
                                f"{prospect.company_name}; no enrichment credits "
                                f"were submitted."
                            )

                        # Parallel Tavily research against the actual ZoomInfo shortlist.
                        worker_count = max(1, min(8, int(self.research_workers.get())))
                        with ThreadPoolExecutor(max_workers=worker_count) as executor:
                            futures = {
                                executor.submit(
                                    research_contact, tavily, contact, prospect.website
                                ): contact
                                for contact in finalists
                            }
                            for future in as_completed(futures):
                                try:
                                    future.result()
                                except Exception as exc:
                                    logger(f"Contact research warning: {exc}")

                        for contact in finalists:
                            # Optional ZoomInfo contact research for CRM relationship and
                            # role detail. This consumes ZoomInfo AI credits.
                            if use_ai_research and "contact_research" in tools:
                                try:
                                    zi_research = mcp.run(mcp.call("contact_research", {
                                        "query": (
                                            "Assess whether this person currently owns or influences "
                                            "marketing, customer acquisition, growth, brand, direct mail, "
                                            "or vendor selection. Also state whether they are already in "
                                            "our CRM or have prior engagement."
                                        ),
                                        "zoominfoContactId": int(contact.contact_id),
                                    }))
                                    zi_text = text_blob(zi_research)
                                    if re.search(
                                        r"already (?:in|exists in) (?:our )?(?:crm|hubspot)|"
                                        r"prior engagement|existing crm record",
                                        zi_text,
                                        re.I,
                                    ):
                                        contact.crm_excluded = True
                                        crm_contact_exclusions += 1
                                    else:
                                        contact.contact_score += 10
                                        contact.live_research_summary += (
                                            " ZoomInfo role/relationship research completed."
                                        )
                                    contact.research_sources += (
                                        " | ZoomInfo contact research"
                                    )
                                except Exception as exc:
                                    logger(f"ZoomInfo contact research warning: {exc}")

                        finalists = [c for c in finalists if not c.crm_excluded]

                        if self.deep_contact_recovery.get() and finalists:
                            preliminary = assign_outreach_order(finalists)
                            recovery_limit = max(
                                1, min(
                                    len(preliminary),
                                    int(self.deep_recovery_contact_limit.get()),
                                )
                            )
                            targets = preliminary[:recovery_limit]
                            with ThreadPoolExecutor(
                                max_workers=min(worker_count, recovery_limit)
                            ) as executor:
                                futures = {
                                    executor.submit(
                                        deep_contact_data_recovery,
                                        tavily, prospect, contact,
                                        preliminary, logger,
                                    ): contact
                                    for contact in targets
                                }
                                for future in as_completed(futures):
                                    try:
                                        future.result()
                                    except Exception as exc:
                                        target_contact = futures[future]
                                        logger(
                                            f"Deep contact recovery warning for "
                                            f"{target_contact.first_name} "
                                            f"{target_contact.last_name}: {exc}"
                                        )

                        finalists = assign_outreach_order(finalists)
                        chosen = finalists[:contact_limit]
                        for index, contact in enumerate(chosen):
                            contact.rank = (
                                "Primary" if index == 0
                                else "Secondary" if index == 1
                                else "Third" if index == 2
                                else f"Alternate {index + 1}"
                            )
                            contact.outreach_order = index + 1
                            contact.outreach_order_label = (
                                "Primary — Contact First" if index == 0
                                else "Secondary — Contact Second" if index == 1
                                else "Third — Contact Third" if index == 2
                                else f"Alternate Contact #{index + 1}"
                            )

                            available_parts = []
                            if contact.email:
                                available_parts.append("email")
                            if contact.direct_phone:
                                available_parts.append("direct phone")
                            if contact.mobile_phone:
                                available_parts.append("mobile")
                            if available_parts:
                                contact.contact_data_status = (
                                    "Available without additional enrichment: "
                                    + ", ".join(available_parts)
                                    + f". Email status: "
                                    f"{contact.email_status or 'unknown'}."
                                )
                            elif contact.public_company_phone:
                                contact.contact_data_status = (
                                    "No direct email or phone found. "
                                    "A public company phone is available."
                                )
                            else:
                                contact.contact_data_status = (
                                    "No public or ZoomInfo contact details found. "
                                    "Paid enrichment was not used."
                                )

                            if contact.decision_maker_confidence >= 88:
                                contact.recommendation_confidence = "High"
                            elif contact.decision_maker_confidence >= 72:
                                contact.recommendation_confidence = "Medium"
                            else:
                                contact.recommendation_confidence = "Fallback"

                            meta = recommendation_meta.get(contact.contact_id, "")
                            if meta:
                                contact.recommendation_reason += f". ZoomInfo recommendation context: {meta}"
                            contact.recommendation_reason += (
                                f". Selected as {contact.rank.lower()} after ranking "
                                f"ZoomInfo and public decision-maker candidates together "
                                f"using role relevance, source quality, public evidence, "
                                f"contactability, and decision-maker confidence."
                            )
                            ranked_contacts.append(contact)
                            self.events.put(("funnel", {
                                "contacts": len(ranked_contacts),
                            }))
                            if contact.rank == "Primary":
                                self.events.put(("prospect_preview", {
                                    "company": prospect.company_name,
                                    "phase": "CONTACT RANKING",
                                    "contact": (
                                        f"{contact.first_name} {contact.last_name} — "
                                        f"{contact.title}. "
                                        f"{contact.recommendation_reason}"
                                    ),
                                    "detail": (
                                        "Company qualified. Compass ranked the "
                                        "best decision-maker using title relevance, "
                                        "source quality, evidence, contactability, "
                                        "and decision-maker confidence."
                                    ),
                                }))
                            history.upsert_master_contact(
                                contact,
                                master_prospect_key,
                            )
                            outreach_drafts.append(
                                generate_outreach_draft(
                                    prospect,
                                    contact,
                                    self.darwill_knowledge,
                                )
                            )

                        self.events.put(("metrics", {
                            "contacts": len(ranked_contacts),
                            "qualified": len(qualified),
                            "reviewed": candidates_reviewed,
                        }))
                        self._save_checkpoint(
                            trade, page, qualified, ranked_contacts, seen_company_ids
                        )

                    if len(records) < 100:
                        break

            summary = {
                "run_time": now_iso(),
                "qualified_companies": len(qualified),
                "requested_companies": target_count,
                "ranked_contacts": len(ranked_contacts),
                "qualification_decisions": len(qualification_decisions),
                "accepted_decisions": sum(1 for item in qualification_decisions if item.get("decision") == "Accepted"),
                "rejected_decisions": sum(1 for item in qualification_decisions if item.get("decision") == "Rejected"),
                "outreach_drafts": len(outreach_drafts),
                "candidates_reviewed": candidates_reviewed,
                "structured_search_pages": pages_called,
                "crm_company_exclusions": crm_company_exclusions,
                "crm_contact_exclusions": crm_contact_exclusions,
                "local_history_exclusions": history_exclusions,
                "master_database_exclusions": master_database_exclusions,
                "hubspot_csv_company_exclusions": hubspot_csv_company_exclusions,
                "hubspot_csv_contact_exclusions": hubspot_csv_contact_exclusions,
                "hubspot_csv_index_enabled": self.use_hubspot_csv_index.get(),
                "master_database_dedup_enabled": self.use_master_dedup.get(),
                "rejected_rereview_days": self.rejected_rereview_days.get(),
                "qualified_rereview_days": self.qualified_rereview_days.get(),
                "search_mode": search_mode,
                "states": states if search_mode == "states" else "Not used",
                "territory_zip": self.territory_zip.get() if search_mode == "radius" else "Not used",
                "territory_radius_miles": self.territory_radius.get() if search_mode == "radius" else "Not used",
                "hard_state_enforcement": search_mode == "states",
                "selected_trade_presets": ", ".join(selected_trade_names),
                "trades": ", ".join(trades),
                "naics_codes": ", ".join(
                    dict.fromkeys(
                        code.strip()
                        for spec in search_specs
                        for code in str(spec.get("naics", "")).split(",")
                        if code.strip()
                    )
                ),
                "adaptive_zoominfo_fallbacks": True,
                "revenue_range": f"${minimum_dollars:,}-${maximum_dollars:,}",
                "zoominfo_revenue_values_sent": f"{minimum_thousands}-{maximum_thousands} (thousands)",
                "employee_range": f"{employee_min}-{employee_max}",
                "crm_ai_checks_enabled": use_ai_research and exclude_crm,
                "contact_enrichment_enabled": enrich_contacts_enabled,
                "smart_email_enrichment_enabled": smart_email_enrichment_enabled,
                "public_decision_maker_discovery": self.discover_public_contacts.get(),
                "deep_contact_recovery": self.deep_contact_recovery.get(),
                "deep_recovery_contact_limit": self.deep_recovery_contact_limit.get(),
                "email_pattern_prediction": self.predict_public_emails.get(),
                "unverified_email_sequence_block": self.block_unverified_sequence_emails.get(),
                "publicly_verified_emails": sum(
                    1 for contact in ranked_contacts
                    if contact.email_verification_status == "Publicly Verified"
                ),
                "predicted_emails_needing_verification": sum(
                    1 for contact in ranked_contacts
                    if contact.email_verification_status == "Needs Verification"
                ),
                "public_contact_data_is_verified": False,
                "predicted_emails_are_verified": False,
                "csv_imports_required": 0,
                "end_reason": (
                    "Requested count reached"
                    if len(qualified) >= target_count
                    else "Structured ZoomInfo search exhausted or stopped"
                ),
            }
            export_workbook(output, qualified, ranked_contacts, outreach_drafts, qualification_decisions, summary)
            self.review_queue = queue_items_from_run(
                qualified,
                ranked_contacts,
                outreach_drafts,
            )
            self.events.put(("refresh_deal_desk", None))
            if self.master_csv_auto_sync.get():
                master_csv_path = self.master_csv_path.get().strip()
                if master_csv_path:
                    try:
                        csv_sync_result = sync_internal_master_to_csv(
                            HistoryDB(DB_PATH),
                            Path(master_csv_path),
                            create_backup=True,
                        )
                        logger(
                            "Permanent master CSV synchronized: "
                            f"{csv_sync_result['records']} records."
                        )
                        self.events.put(
                            ("refresh_master_csv_status", None)
                        )
                    except Exception as exc:
                        logger(
                            f"Permanent master CSV sync warning: {exc}"
                        )

            self.events.put(("refresh_master_database", None))
            self.events.put((
                "hubspot_index_duplicates",
                (
                    hubspot_csv_company_exclusions
                    + hubspot_csv_contact_exclusions
                ),
            ))
            if self.current_run_id is not None:
                history.finish_run(
                    self.current_run_id,
                    len(qualified),
                    len(ranked_contacts),
                    candidates_reviewed,
                    str(output),
                    "COMPLETED" if not self.stop_requested else "STOPPED",
                )
            CHECKPOINT_FILE.unlink(missing_ok=True)
            logger(f"Workbook saved: {output}")
            self.events.put((
                "status",
                f"Completed: {len(qualified)} qualified companies and "
                f"{len(ranked_contacts)} ranked contacts."
            ))
        except Exception as exc:
            try:
                if self.current_run_id is not None:
                    HistoryDB(DB_PATH).finish_run(
                        self.current_run_id, 0, 0, 0, str(output), "FAILED"
                    )
            except Exception:
                pass
            self.events.put(("log", f"ERROR: {exc}\n{traceback.format_exc()}"))
            self.events.put(("status", "Run failed—see activity log and logs folder."))
        finally:
            self.events.put(("done", None))

    def _poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self._append(value)
                elif kind == "status":
                    self.status.set(value)
                    if hasattr(self, "footer_status"):
                        self.footer_status.set(value)
                elif kind == "refresh_deal_desk":
                    self._refresh_deal_desk()
                elif kind == "refresh_master_database":
                    self._refresh_master_database()
                elif kind == "refresh_master_csv_status":
                    self._refresh_master_csv_status()
                elif kind == "hubspot_index_duplicates":
                    if hasattr(self, "hubspot_index_duplicates_prevented"):
                        self.hubspot_index_duplicates_prevented.set(
                            f"{int(value):,}"
                        )
                elif kind == "metrics":
                    for key, metric_value in value.items():
                        if key in self.metric_vars:
                            self.metric_vars[key].set(str(metric_value))
                    if "reviewed" in value:
                        self._update_funnel(
                            "reviewed",
                            int(value["reviewed"]),
                        )
                    if "qualified" in value:
                        self._update_funnel(
                            "qualified",
                            int(value["qualified"]),
                        )
                    if "contacts" in value:
                        self._update_funnel(
                            "contacts",
                            int(value["contacts"]),
                        )
                elif kind == "funnel":
                    for key, count in value.items():
                        self._update_funnel(key, int(count))
                elif kind == "funnel_add":
                    for key, count in value.items():
                        self._update_funnel(
                            key,
                            int(count),
                            add=True,
                        )
                elif kind == "prospect_preview":
                    self._set_prospect_preview(value)
                elif kind == "connected_done":
                    self.progress.stop()
                elif kind == "done":
                    self.progress.stop()
                    self.run_started_at = 0.0
                    self.run_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self._refresh_history()
        except queue.Empty:
            pass
        self.after(100, self._poll)


if __name__ == "__main__":
    App().mainloop()
