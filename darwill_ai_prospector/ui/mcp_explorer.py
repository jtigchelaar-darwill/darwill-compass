"""Standalone ZoomInfo MCP Explorer window."""

from __future__ import annotations

import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Callable


def _schema_for_tool(tool_value: Any) -> dict[str, Any]:
    if isinstance(tool_value, dict):
        return tool_value

    result: dict[str, Any] = {}
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


def open_mcp_explorer_window(
    *,
    parent: tk.Misc,
    app_title: str,
    log_dir: Path,
    connection_factory: Callable[[], Any],
) -> None:
    """Open a self-contained MCP schema and request explorer."""
    explorer_bg = "#F4F6F8"
    explorer_navy = "#0B2D4F"
    explorer_white = "#FFFFFF"
    explorer_muted = "#CFE3F6"

    explorer = tk.Toplevel(parent)
    explorer.title("ZoomInfo MCP Explorer")
    explorer.geometry("1250x780")
    explorer.minsize(980, 620)
    explorer.configure(bg=explorer_bg)
    explorer.transient(parent)

    header = tk.Frame(explorer, bg=explorer_navy, padx=18, pady=14)
    header.pack(fill="x")
    tk.Label(
        header,
        text="ZOOMINFO DEVELOPER TOOLS",
        bg=explorer_navy,
        fg="#75B6F5",
        font=("Segoe UI Semibold", 8),
    ).pack(anchor="w")
    tk.Label(
        header,
        text="MCP Explorer",
        bg=explorer_navy,
        fg=explorer_white,
        font=("Segoe UI Semibold", 20),
    ).pack(anchor="w", pady=(2, 2))
    tk.Label(
        header,
        text=(
            "Inspect live tools and schemas, then run controlled tests "
            "without rerunning Discovery."
        ),
        bg=explorer_navy,
        fg=explorer_muted,
        font=("Segoe UI", 10),
    ).pack(anchor="w")

    toolbar = ttk.Frame(explorer, padding=(14, 12))
    toolbar.pack(fill="x")
    status_var = tk.StringVar(value="Not connected")

    body = ttk.Panedwindow(explorer, orient="horizontal")
    body.pack(fill="both", expand=True, padx=14, pady=(0, 14))

    left = ttk.Frame(body, padding=8)
    middle = ttk.Frame(body, padding=8)
    right = ttk.Frame(body, padding=8)
    body.add(left, weight=1)
    body.add(middle, weight=2)
    body.add(right, weight=2)

    ttk.Label(left, text="Available Tools").pack(anchor="w", pady=(0, 6))
    tool_filter_var = tk.StringVar()
    ttk.Entry(left, textvariable=tool_filter_var).pack(
        fill="x",
        pady=(0, 6),
    )
    tool_list = tk.Listbox(
        left,
        exportselection=False,
        font=("Consolas", 10),
    )
    tool_list.pack(fill="both", expand=True)

    ttk.Label(middle, text="Tool Schema").pack(anchor="w", pady=(0, 6))
    schema_text = tk.Text(middle, wrap="none", font=("Consolas", 9))
    schema_text.pack(fill="both", expand=True)

    ttk.Label(right, text="Request and Response").pack(
        anchor="w",
        pady=(0, 6),
    )
    request_frame = ttk.LabelFrame(right, text="Request JSON", padding=6)
    request_frame.pack(fill="both", expand=True)
    request_text = tk.Text(
        request_frame,
        wrap="none",
        font=("Consolas", 9),
        height=13,
    )
    request_text.pack(fill="both", expand=True)

    action_bar = ttk.Frame(right)
    action_bar.pack(fill="x", pady=8)

    response_frame = ttk.LabelFrame(right, text="Raw Response", padding=6)
    response_frame.pack(fill="both", expand=True)
    response_text = tk.Text(
        response_frame,
        wrap="none",
        font=("Consolas", 9),
        height=13,
    )
    response_text.pack(fill="both", expand=True)

    state: dict[str, Any] = {
        "mcp": None,
        "tools": {},
        "visible_names": [],
        "last_request": {},
        "last_response": {},
        "selected_tool": "",
    }

    def pretty(value: Any) -> str:
        return json.dumps(value, indent=2, default=str, sort_keys=True)

    def selected_tool_name() -> str:
        selection = tool_list.curselection()
        if not selection:
            return ""
        index = selection[0]
        names = state["visible_names"]
        return names[index] if index < len(names) else ""

    def refresh_list(*_args: Any) -> None:
        query = tool_filter_var.get().strip().lower()
        names = sorted(state["tools"].keys())
        if query:
            names = [name for name in names if query in name.lower()]
        state["visible_names"] = names
        tool_list.delete(0, "end")
        for name in names:
            tool_list.insert("end", name)
        status_var.set(
            f"{len(names)} shown / {len(state['tools'])} tools loaded"
            if state["tools"]
            else "Not connected"
        )

    def show_schema(_event: Any = None) -> None:
        name = selected_tool_name()
        if not name:
            return
        state["selected_tool"] = name
        schema = _schema_for_tool(state["tools"][name])
        schema_text.delete("1.0", "end")
        schema_text.insert("1.0", pretty(schema))

    def load_tools() -> None:
        try:
            status_var.set("Connecting to ZoomInfo MCP…")
            explorer.update_idletasks()
            mcp = connection_factory()
            tools = mcp.run(mcp.discover())
            if not isinstance(tools, dict):
                raise RuntimeError(
                    "MCP discovery returned an unexpected tool collection."
                )
            state["mcp"] = mcp
            state["tools"] = tools
            refresh_list()
            if state["visible_names"]:
                tool_list.selection_set(0)
                show_schema()
            status_var.set(f"{len(tools)} MCP tools loaded")
        except Exception as exc:
            status_var.set("Connection failed")
            messagebox.showerror(
                app_title,
                f"Could not load MCP tools:\n\n{exc}",
                parent=explorer,
            )

    def schema_properties(name: str) -> dict[str, Any]:
        tool = _schema_for_tool(state["tools"].get(name, {}))
        schema = (
            tool.get("inputSchema")
            or tool.get("input_schema")
            or tool.get("schema")
            or {}
        )
        return schema if isinstance(schema, dict) else {}

    def load_sample_request() -> None:
        name = selected_tool_name()
        if not name:
            messagebox.showinfo(
                app_title,
                "Select a tool first.",
                parent=explorer,
            )
            return

        schema = schema_properties(name)
        props = schema.get("properties", {})
        required = schema.get("required", [])
        sample: dict[str, Any] = {}

        def placeholder(field: str, definition: dict[str, Any]) -> Any:
            field_lower = field.lower()
            if field_lower in {"personid", "contactid"}:
                return "ZOOMINFO_PERSON_ID"
            if field_lower in {"companyid", "zoominfocompanyid"}:
                return "ZOOMINFO_COMPANY_ID"
            if field_lower == "firstname":
                return "Gabe"
            if field_lower == "lastname":
                return "Garwick"
            if field_lower in {"companyname", "company"}:
                return "Mr. Electric"
            if field_lower in {"jobtitle", "title"}:
                return "Marketing Specialist"
            if field_lower == "requiredfields":
                return [
                    "firstName",
                    "lastName",
                    "email",
                    "phone",
                    "mobilePhone",
                    "jobTitle",
                    "companyName",
                ]
            if field_lower == "contacts":
                return [
                    {
                        "firstName": "Gabe",
                        "lastName": "Garwick",
                        "companyName": "Mr. Electric",
                        "jobTitle": "Marketing Specialist",
                    }
                ]
            field_type = definition.get("type")
            if field_type == "boolean":
                return True
            if field_type in {"integer", "number"}:
                return 1
            if field_type == "array":
                return []
            if field_type == "object":
                return {}
            return ""

        for field in required:
            sample[field] = placeholder(field, props.get(field, {}))

        if name == "enrich_contacts":
            sample = {
                "contacts": [
                    {
                        "firstName": "Gabe",
                        "lastName": "Garwick",
                        "companyName": "Mr. Electric",
                        "jobTitle": "Marketing Specialist",
                    }
                ],
                "requiredFields": [
                    "firstName",
                    "lastName",
                    "email",
                    "phone",
                    "mobilePhone",
                    "jobTitle",
                    "companyName",
                ],
                "userIntent": (
                    "Test one contact and return the verified business "
                    "email when available."
                ),
            }
        elif name == "search_contacts":
            sample.update(
                {
                    "firstName": "Gabe",
                    "lastName": "Garwick",
                    "companyName": "Mr. Electric",
                }
            )

        request_text.delete("1.0", "end")
        request_text.insert("1.0", pretty(sample))

    def run_selected_tool() -> None:
        name = selected_tool_name()
        if not name:
            messagebox.showinfo(
                app_title,
                "Select a tool first.",
                parent=explorer,
            )
            return
        if state["mcp"] is None:
            messagebox.showinfo(
                app_title,
                "Connect and load tools first.",
                parent=explorer,
            )
            return

        try:
            payload = json.loads(
                request_text.get("1.0", "end").strip() or "{}"
            )
        except json.JSONDecodeError as exc:
            messagebox.showerror(
                app_title,
                f"Request JSON is invalid:\n\n{exc}",
                parent=explorer,
            )
            return

        try:
            status_var.set(f"Running {name}…")
            explorer.update_idletasks()
            response = state["mcp"].run(
                state["mcp"].call(name, payload)
            )
            state["last_request"] = payload
            state["last_response"] = response
            state["selected_tool"] = name
            response_text.delete("1.0", "end")
            response_text.insert("1.0", pretty(response))

            session_dir = Path(log_dir) / "mcp_explorer"
            session_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            tool_dir = session_dir / f"{stamp}_{name}"
            tool_dir.mkdir(parents=True, exist_ok=True)
            (tool_dir / "tool_schema.json").write_text(
                schema_text.get("1.0", "end").strip(),
                encoding="utf-8",
            )
            (tool_dir / "request.json").write_text(
                pretty(payload),
                encoding="utf-8",
            )
            (tool_dir / "response.json").write_text(
                pretty(response),
                encoding="utf-8",
            )
            status_var.set(f"{name} completed — saved to {tool_dir}")
        except Exception as exc:
            status_var.set(f"{name} failed")
            response_text.delete("1.0", "end")
            response_text.insert(
                "1.0",
                pretty(
                    {
                        "success": False,
                        "tool": name,
                        "error": str(exc),
                    }
                ),
            )
            messagebox.showerror(
                app_title,
                f"MCP tool failed:\n\n{exc}",
                parent=explorer,
            )

    def save_session() -> None:
        session_dir = Path(log_dir) / "mcp_explorer"
        session_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = session_dir / f"{stamp}_manual_session.json"
        output.write_text(
            pretty(
                {
                    "selectedTool": state["selected_tool"],
                    "tools": {
                        name: _schema_for_tool(value)
                        for name, value in state["tools"].items()
                    },
                    "lastRequest": state["last_request"],
                    "lastResponse": state["last_response"],
                }
            ),
            encoding="utf-8",
        )
        messagebox.showinfo(
            app_title,
            f"Explorer session saved to:\n\n{output}",
            parent=explorer,
        )

    ttk.Button(
        toolbar,
        text="Connect and Load Tools",
        command=load_tools,
    ).pack(side="left")
    ttk.Button(
        toolbar,
        text="Save Explorer Session",
        command=save_session,
    ).pack(side="left", padx=(8, 0))
    ttk.Label(toolbar, textvariable=status_var).pack(
        side="left",
        padx=(12, 0),
    )

    ttk.Button(
        action_bar,
        text="Load Sample Request",
        command=load_sample_request,
    ).pack(side="left")
    ttk.Button(
        action_bar,
        text="Run Selected Tool",
        command=run_selected_tool,
    ).pack(side="left", padx=(8, 0))

    tool_filter_var.trace_add("write", refresh_list)
    tool_list.bind("<<ListboxSelect>>", show_schema)

    request_text.insert(
        "1.0",
        pretty(
            {
                "instructions": (
                    "Connect, select a tool, then load or enter a request."
                )
            }
        ),
    )
    response_text.insert(
        "1.0",
        pretty({"status": "No tool has been run yet."}),
    )
    explorer.update_idletasks()
    explorer.lift()
    explorer.focus_force()
