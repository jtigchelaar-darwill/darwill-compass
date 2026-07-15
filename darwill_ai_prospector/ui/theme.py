"""Shared Compass 6.0 workspace theme."""

from __future__ import annotations

from tkinter import ttk


def workspace_palette() -> dict[str, str]:
    return {
        "command_bg": "#F6F9FC",
        "content_bg": "#EEF3F8",
        "surface": "#FFFFFF",
        "border": "#D8E4EF",
        "text": "#17324D",
        "muted": "#5A7184",
        "accent": "#0B6FD3",
        "success": "#16794C",
    }


def apply_compass_theme(style: ttk.Style) -> None:
    """Standardize shared ttk controls without changing app behavior."""
    try:
        style.theme_use("clam")
    except Exception:
        pass

    palette = workspace_palette()

    style.configure(
        "Workspace.TFrame",
        background=palette["content_bg"],
    )
    style.configure(
        "Card.TFrame",
        background=palette["surface"],
    )
    style.configure(
        "Card.TLabel",
        background=palette["surface"],
        foreground=palette["text"],
        font=("Segoe UI", 9),
    )
    style.configure(
        "SectionTitle.TLabel",
        background=palette["surface"],
        foreground=palette["text"],
        font=("Segoe UI Semibold", 11),
    )
    style.configure(
        "Muted.TLabel",
        background=palette["surface"],
        foreground=palette["muted"],
        font=("Segoe UI", 8),
    )
