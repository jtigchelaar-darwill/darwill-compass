"""Centralized ZoomInfo MCP access, logging, and circuit breaking."""

from __future__ import annotations

import json
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


@dataclass(frozen=True)
class ZoomInfoCallResult:
    tool_name: str
    request: dict[str, Any]
    response: Any
    classification: str
    success: bool
    log_dir: Path


class ZoomInfoAdapterError(RuntimeError):
    pass


class ZoomInfoAdapter:
    def __init__(
        self,
        mcp: Any,
        *,
        log_root: Path,
        logger: Callable[[str], None] | None = None,
    ):
        self.mcp = mcp
        self.log_root = Path(log_root) / "zoominfo_adapter"
        self.log_root.mkdir(parents=True, exist_ok=True)
        self.logger = logger or (lambda _message: None)
        self.tools: dict[str, Any] = {}
        self.enrichment_blocked = False
        self.enrichment_block_reason = ""

    @staticmethod
    def _blob(value: Any) -> str:
        try:
            return json.dumps(value, default=str).lower()
        except Exception:
            return str(value).lower()

    @classmethod
    def classify(cls, response: Any) -> str:
        blob = cls._blob(response)
        if "limit exceeded" in blob:
            return "limit_exceeded"
        if any(
            phrase in blob
            for phrase in [
                "not entitled",
                "permission denied",
                "forbidden",
                "unauthorized",
                "entitlement",
            ]
        ):
            return "permission_or_entitlement"
        if any(
            phrase in blob
            for phrase in [
                "rate limit",
                "too many requests",
                "throttl",
            ]
        ):
            return "rate_limited"
        if any(
            phrase in blob
            for phrase in [
                '"success": false',
                "'success': false",
                '"totalerrors":',
            ]
        ):
            return "tool_error"
        return "success"

    def discover_tools(self, *, force: bool = False) -> dict[str, Any]:
        if self.tools and not force:
            return self.tools

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        call_dir = self.log_root / f"{stamp}_discover_tools"
        call_dir.mkdir(parents=True, exist_ok=True)
        response = self.mcp.run(self.mcp.discover())
        (call_dir / "response.json").write_text(
            json.dumps(response, indent=2, default=str),
            encoding="utf-8",
        )
        if not isinstance(response, dict):
            raise ZoomInfoAdapterError(
                "ZoomInfo MCP tool discovery returned an unexpected response."
            )
        self.tools = response
        return response

    def has_tool(self, name: str) -> bool:
        return name in self.discover_tools()

    def call(
        self,
        tool_name: str,
        payload: dict[str, Any],
        *,
        retry_count: int = 1,
        block_on_limit: bool = True,
    ) -> ZoomInfoCallResult:
        if tool_name == "enrich_contacts" and self.enrichment_blocked:
            raise ZoomInfoAdapterError(
                "ZoomInfo enrichment is disabled for this session: "
                + self.enrichment_block_reason
            )

        tools = self.discover_tools()
        if tool_name not in tools:
            raise ZoomInfoAdapterError(
                f"ZoomInfo MCP does not expose {tool_name}."
            )

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        call_dir = self.log_root / f"{stamp}_{tool_name}"
        call_dir.mkdir(parents=True, exist_ok=True)
        (call_dir / "request.json").write_text(
            json.dumps(payload, indent=2, default=str),
            encoding="utf-8",
        )
        (call_dir / "tool_schema.json").write_text(
            json.dumps(tools.get(tool_name, {}), indent=2, default=str),
            encoding="utf-8",
        )

        response: Any = None
        attempts = max(1, retry_count + 1)
        for attempt in range(1, attempts + 1):
            try:
                response = self.mcp.run(self.mcp.call(tool_name, payload))
                break
            except Exception as exc:
                if attempt >= attempts:
                    (call_dir / "exception.txt").write_text(
                        str(exc),
                        encoding="utf-8",
                    )
                    raise
                time.sleep(min(2 ** (attempt - 1), 3))

        (call_dir / "response.json").write_text(
            json.dumps(response, indent=2, default=str),
            encoding="utf-8",
        )
        classification = self.classify(response)
        success = classification == "success"
        (call_dir / "summary.json").write_text(
            json.dumps(
                {
                    "tool": tool_name,
                    "classification": classification,
                    "success": success,
                    "attempts": attempt,
                },
                indent=2,
            ),
            encoding="utf-8",
        )

        if (
            tool_name == "enrich_contacts"
            and classification == "limit_exceeded"
            and block_on_limit
        ):
            self.enrichment_blocked = True
            self.enrichment_block_reason = (
                "ZoomInfo MCP returned 'Limit exceeded'. Further enrichment "
                "calls are skipped until Compass restarts."
            )
            self.logger(self.enrichment_block_reason)

        return ZoomInfoCallResult(
            tool_name=tool_name,
            request=payload,
            response=response,
            classification=classification,
            success=success,
            log_dir=call_dir,
        )

    def reset_enrichment_block(self) -> None:
        self.enrichment_blocked = False
        self.enrichment_block_reason = ""
