"""Reusable scrolling helpers for Darwill Compass Tkinter widgets."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


def _bind_mousewheel(widget: tk.Widget, *, horizontal: bool = False) -> None:
    """Bind Windows/macOS/Linux mouse-wheel events while pointer is inside."""

    def on_mousewheel(event: tk.Event) -> str:
        delta = getattr(event, "delta", 0)
        if delta:
            units = -1 * int(delta / 120)
        else:
            button = getattr(event, "num", 0)
            units = -1 if button == 4 else 1

        if horizontal:
            widget.xview_scroll(units, "units")
        else:
            widget.yview_scroll(units, "units")
        return "break"

    def bind_events(_event: tk.Event) -> None:
        widget.bind_all("<MouseWheel>", on_mousewheel)
        widget.bind_all("<Button-4>", on_mousewheel)
        widget.bind_all("<Button-5>", on_mousewheel)

    def unbind_events(_event: tk.Event) -> None:
        widget.unbind_all("<MouseWheel>")
        widget.unbind_all("<Button-4>")
        widget.unbind_all("<Button-5>")

    widget.bind("<Enter>", bind_events, add="+")
    widget.bind("<Leave>", unbind_events, add="+")


def configure_scrollable_tree(
    parent: tk.Widget,
    tree: ttk.Treeview,
    *,
    horizontal: bool = True,
    vertical: bool = True,
    enable_mousewheel: bool = True,
) -> tuple[ttk.Scrollbar | None, ttk.Scrollbar | None]:
    """Lay out a Treeview with optional horizontal and vertical scrollbars."""
    parent.rowconfigure(0, weight=1)
    parent.columnconfigure(0, weight=1)

    y_scroll = None
    x_scroll = None

    if vertical:
        y_scroll = ttk.Scrollbar(
            parent,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=y_scroll.set)
        y_scroll.grid(row=0, column=1, sticky="ns")

    if horizontal:
        x_scroll = ttk.Scrollbar(
            parent,
            orient="horizontal",
            command=tree.xview,
        )
        tree.configure(xscrollcommand=x_scroll.set)
        x_scroll.grid(row=1, column=0, sticky="ew")

    tree.grid(row=0, column=0, sticky="nsew")

    if enable_mousewheel:
        _bind_mousewheel(tree)

    return y_scroll, x_scroll


def configure_scrollable_text(
    parent: tk.Widget,
    text_widget: tk.Text,
    *,
    horizontal: bool = False,
    vertical: bool = True,
    enable_mousewheel: bool = True,
) -> tuple[ttk.Scrollbar | None, ttk.Scrollbar | None]:
    """Lay out a Text widget with optional scrollbars and wheel support."""
    parent.rowconfigure(0, weight=1)
    parent.columnconfigure(0, weight=1)

    y_scroll = None
    x_scroll = None

    if vertical:
        y_scroll = ttk.Scrollbar(
            parent,
            orient="vertical",
            command=text_widget.yview,
        )
        text_widget.configure(yscrollcommand=y_scroll.set)
        y_scroll.grid(row=0, column=1, sticky="ns")

    if horizontal:
        x_scroll = ttk.Scrollbar(
            parent,
            orient="horizontal",
            command=text_widget.xview,
        )
        text_widget.configure(xscrollcommand=x_scroll.set)
        x_scroll.grid(row=1, column=0, sticky="ew")

    text_widget.grid(row=0, column=0, sticky="nsew")

    if enable_mousewheel:
        _bind_mousewheel(text_widget)

    return y_scroll, x_scroll



class VerticalScrolledFrame(ttk.Frame):
    """A vertically scrollable frame with a normal child content frame.

    Widgets should be created with ``instance.content`` as their parent.
    The content width follows the viewport while its height may grow freely.
    """

    def __init__(
        self,
        parent: tk.Widget,
        *,
        background: str = "#FFFFFF",
        padding: int = 0,
    ) -> None:
        super().__init__(parent)

        self.canvas = tk.Canvas(
            self,
            borderwidth=0,
            highlightthickness=0,
            background=background,
        )
        self.scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.canvas.yview,
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.content = ttk.Frame(
            self.canvas,
            padding=padding,
            style="Card.TFrame",
        )
        self._window_id = self.canvas.create_window(
            (0, 0),
            window=self.content,
            anchor="nw",
        )

        self.content.bind(
            "<Configure>",
            self._sync_scroll_region,
            add="+",
        )
        self.canvas.bind(
            "<Configure>",
            self._sync_content_width,
            add="+",
        )
        _bind_mousewheel(self.canvas)

    def _sync_scroll_region(self, _event: tk.Event | None = None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_content_width(self, event: tk.Event) -> None:
        self.canvas.itemconfigure(self._window_id, width=event.width)
        self._sync_scroll_region()

    def scroll_to_top(self) -> None:
        self.canvas.yview_moveto(0.0)
