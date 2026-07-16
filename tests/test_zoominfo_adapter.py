from pathlib import Path

from darwill_ai_prospector.services.zoominfo_adapter import ZoomInfoAdapter


class FakeMCP:
    def __init__(self, response):
        self.response = response

    def discover(self):
        return ("discover", {})

    def call(self, name, payload):
        return (name, payload)

    def run(self, operation):
        if operation[0] == "discover":
            return {"enrich_contacts": {}}
        return self.response


def test_limit_exceeded_blocks_future_calls(tmp_path: Path):
    adapter = ZoomInfoAdapter(
        FakeMCP({"error": "Contact enrichment failed: Limit exceeded"}),
        log_root=tmp_path,
    )
    result = adapter.call(
        "enrich_contacts",
        {"contacts": []},
        retry_count=0,
    )
    assert result.classification == "limit_exceeded"
    assert adapter.enrichment_blocked is True
