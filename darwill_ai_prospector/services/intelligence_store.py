"""SQLite evidence, timeline, and email-intelligence storage."""

from __future__ import annotations
import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _get(item: Any, name: str, default: Any = "") -> Any:
    return getattr(item, name, default)


class IntelligenceStore:
    def __init__(self, path: Path):
        self.conn = sqlite3.connect(Path(path))
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._migrate()

    def _migrate(self) -> None:
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS intelligence_prospects(
            prospect_key TEXT PRIMARY KEY,
            company_id TEXT,
            company_name TEXT NOT NULL,
            website TEXT,
            city TEXT,
            state TEXT,
            revenue INTEGER,
            employees INTEGER,
            lifecycle_status TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS intelligence_contacts(
            contact_key TEXT PRIMARY KEY,
            prospect_key TEXT NOT NULL,
            contact_id TEXT,
            full_name TEXT,
            title TEXT,
            email TEXT,
            phone TEXT,
            decision_confidence INTEGER DEFAULT 0,
            email_confidence INTEGER DEFAULT 0,
            phone_confidence INTEGER DEFAULT 0,
            source_type TEXT,
            source_url TEXT,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS intelligence_evidence(
            evidence_id INTEGER PRIMARY KEY AUTOINCREMENT,
            prospect_key TEXT NOT NULL,
            contact_key TEXT,
            category TEXT NOT NULL,
            finding TEXT NOT NULL,
            value TEXT,
            source_type TEXT,
            source_url TEXT,
            confidence INTEGER DEFAULT 0,
            verification_status TEXT,
            observed_at TEXT NOT NULL,
            UNIQUE(prospect_key, contact_key, category, finding, value, source_url)
        );
        CREATE TABLE IF NOT EXISTS email_intelligence(
            email_intelligence_id INTEGER PRIMARY KEY AUTOINCREMENT,
            prospect_key TEXT NOT NULL,
            contact_key TEXT NOT NULL,
            email TEXT,
            source_type TEXT,
            public_found INTEGER DEFAULT 0,
            predicted_pattern TEXT,
            pattern_support INTEGER DEFAULT 0,
            confidence INTEGER DEFAULT 0,
            verification_status TEXT,
            recovery_method TEXT,
            zoominfo_status TEXT,
            credit_recommended INTEGER DEFAULT 0,
            recommendation TEXT,
            expected_gain TEXT,
            evaluated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS intelligence_timeline(
            timeline_id INTEGER PRIMARY KEY AUTOINCREMENT,
            prospect_key TEXT NOT NULL,
            event_type TEXT NOT NULL,
            event_label TEXT NOT NULL,
            engine TEXT,
            details_json TEXT,
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_intel_evidence_prospect
            ON intelligence_evidence(prospect_key);
        CREATE INDEX IF NOT EXISTS idx_intel_contacts_prospect
            ON intelligence_contacts(prospect_key);
        CREATE INDEX IF NOT EXISTS idx_intel_timeline_prospect
            ON intelligence_timeline(prospect_key, created_at);
        """)
        self.conn.commit()

    @staticmethod
    def prospect_key(item: Any) -> str:
        raw = (
            str(_get(item, "company_id", "") or "").strip()
            or str(_get(item, "company_website", "") or "").strip().lower()
            or str(_get(item, "company_name", "") or "").strip().lower()
        )
        return "prospect-" + hashlib.sha1(raw.encode()).hexdigest()[:20]

    @staticmethod
    def contact_key(item: Any) -> str:
        raw = (
            str(_get(item, "contact_id", "") or "").strip()
            or f"{_get(item, 'contact_name', '')}|{_get(item, 'contact_title', '')}".lower()
        )
        return "contact-" + hashlib.sha1(raw.encode()).hexdigest()[:20]

    def upsert_queue_item(self, item: Any) -> str:
        pkey, ckey, now = self.prospect_key(item), self.contact_key(item), _now()
        self.conn.execute("""
            INSERT INTO intelligence_prospects(
                prospect_key, company_id, company_name, website, city, state,
                revenue, employees, lifecycle_status, created_at, updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(prospect_key) DO UPDATE SET
                company_name=excluded.company_name,
                website=excluded.website, city=excluded.city,
                state=excluded.state, revenue=excluded.revenue,
                employees=excluded.employees,
                lifecycle_status=excluded.lifecycle_status,
                updated_at=excluded.updated_at
        """, (
            pkey, _get(item,"company_id",""), _get(item,"company_name",""),
            _get(item,"company_website",""), _get(item,"company_city",""),
            _get(item,"company_state",""), _get(item,"company_revenue",None),
            _get(item,"company_employees",None), _get(item,"status",""),
            now, now,
        ))
        self.conn.execute("""
            INSERT INTO intelligence_contacts(
                contact_key, prospect_key, contact_id, full_name, title,
                email, phone, decision_confidence, email_confidence,
                phone_confidence, source_type, source_url, updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(contact_key) DO UPDATE SET
                prospect_key=excluded.prospect_key,
                full_name=excluded.full_name, title=excluded.title,
                email=excluded.email, phone=excluded.phone,
                decision_confidence=excluded.decision_confidence,
                email_confidence=excluded.email_confidence,
                phone_confidence=excluded.phone_confidence,
                source_type=excluded.source_type,
                source_url=excluded.source_url, updated_at=excluded.updated_at
        """, (
            ckey, pkey, _get(item,"contact_id",""),
            _get(item,"contact_name",""), _get(item,"contact_title",""),
            _get(item,"contact_email",""), _get(item,"contact_phone",""),
            int(_get(item,"contact_decision_confidence",0) or 0),
            int(_get(item,"contact_email_confidence",0) or 0),
            int(_get(item,"contact_phone_confidence",0) or 0),
            _get(item,"contact_source_type",""),
            _get(item,"contact_source_url",""), now,
        ))
        self.conn.commit()
        return pkey

    def add_evidence(self, pkey, category, finding, value, *,
                     contact_key="", source_type="", source_url="",
                     confidence=0, verification_status=""):
        self.conn.execute("""
            INSERT OR IGNORE INTO intelligence_evidence(
                prospect_key, contact_key, category, finding, value,
                source_type, source_url, confidence, verification_status,
                observed_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?)
        """, (
            pkey, contact_key or None, category, finding, str(value or ""),
            source_type, source_url, int(confidence or 0),
            verification_status, _now(),
        ))
        self.conn.commit()

    def record_standard_evidence(self, pkey, item, decision):
        ckey = self.contact_key(item)
        source_type = str(_get(item,"contact_source_type","") or "")
        source_url = str(
            _get(item,"contact_research_sources","")
            or _get(item,"contact_source_url","")
            or _get(item,"sources","") or ""
        )
        rows = [
            ("qualification","why_qualified",_get(item,"why_company",""),
             int(_get(item,"outreach_confidence",0) or 0),"stored",""),
            ("contact","recommended_contact",
             f"{_get(item,'contact_name','')} — {_get(item,'contact_title','')}",
             int(_get(item,"contact_decision_confidence",0) or 0),"ranked",ckey),
            ("email","email_address",_get(item,"contact_email","") or "Not found",
             int(_get(item,"contact_email_confidence",0) or 0),
             _get(item,"contact_email_verification_status",""),ckey),
            ("email","credit_recommendation",decision.recommendation,
             decision.confidence,
             "recommended" if decision.credit_recommended else "not_recommended",ckey),
            ("opportunity","recommended_strategy",
             _get(item,"recommended_strategy",""),
             int(_get(item,"outreach_confidence",0) or 0),"stored",""),
        ]
        for cat, finding, value, conf, verify, contact in rows:
            self.add_evidence(
                pkey, cat, finding, value, contact_key=contact,
                source_type=source_type, source_url=source_url,
                confidence=conf, verification_status=verify,
            )
        self.conn.execute("""
            INSERT INTO email_intelligence(
                prospect_key, contact_key, email, source_type, public_found,
                predicted_pattern, pattern_support, confidence,
                verification_status, recovery_method, zoominfo_status,
                credit_recommended, recommendation, expected_gain, evaluated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            pkey, ckey, _get(item,"contact_email",""), source_type,
            1 if decision.public_status == "Found free" else 0,
            _get(item,"contact_predicted_email_pattern",""),
            int(_get(item,"contact_pattern_support_count",0) or 0),
            decision.confidence,
            _get(item,"contact_email_verification_status",""),
            _get(item,"contact_email_recovery_method",""),
            decision.zoominfo_status,
            1 if decision.credit_recommended else 0,
            decision.recommendation, decision.expected_gain, _now(),
        ))
        self.conn.commit()

    def append_timeline(self, pkey, event_type, event_label, engine, details=None):
        self.conn.execute("""
            INSERT INTO intelligence_timeline(
                prospect_key, event_type, event_label, engine,
                details_json, created_at
            ) VALUES(?,?,?,?,?,?)
        """, (
            pkey, event_type, event_label, engine,
            json.dumps(details or {}, sort_keys=True), _now(),
        ))
        self.conn.commit()
