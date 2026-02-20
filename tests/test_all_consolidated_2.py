from __future__ import annotations

import tkinter as tk
from tkinter import ttk

import pytest


def _find_button_by_text(root: tk.Misc, text: str) -> ttk.Button | None:
    for w in root.winfo_children():
        if isinstance(w, ttk.Button) and w.cget("text") == text:
            return w
        # Recurse into frames
        if isinstance(w, (tk.Frame, ttk.Frame, ttk.LabelFrame)):
            found = _find_button_by_text(w, text)
            if found is not None:
                return found
    return None


def _topbar_buttons_in_grid_order(app) -> list[str]:
    # Find the frame that contains the topbar buttons by locating the Add Source button.
    btn = _find_button_by_text(app, "Add Source (XLSX/CSV)")
    assert btn is not None, "Add Source button not found"
    topbar = btn.master
    # Collect row 0 slaves and sort by column.
    slaves = list(topbar.grid_slaves(row=0))
    def col(w):
        info = w.grid_info()
        return int(info.get("column", 0))
    slaves.sort(key=col)
    return [w.cget("text") for w in slaves if isinstance(w, ttk.Button)]


def test_add_source_button_matches_run_button_style():
    from gui.app import TurboExtractorApp

    app = TurboExtractorApp()
    try:
        btn_add = _find_button_by_text(app, "Add Source (XLSX/CSV)")
        btn_run = _find_button_by_text(app, "RUN")
        assert btn_add is not None and btn_run is not None
        assert btn_add.cget("style") == btn_run.cget("style")
    finally:
        app.destroy()


def test_topbar_button_order_matches_spec():
    from gui.app import TurboExtractorApp

    app = TurboExtractorApp()
    try:
        order = _topbar_buttons_in_grid_order(app)
        assert order[:6] == [
            "Add Source (XLSX/CSV)",
            "Move Source Up",
            "Move Source Down",
            "Add Recipe",
            "Add Sheet",
            "Remove Selected",
        ]
    finally:
        app.destroy()


def test_tree_is_fully_expanded_by_default():
    from gui.app import TurboExtractorApp
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app = TurboExtractorApp()
    try:
        # Add a source by manipulating the project model directly (no dialogs)
        app.project.sources.append(
            SourceConfig(
                path="C:/tmp/example.xlsx",
                recipes=[
                    RecipeConfig(
                        name="Recipe1",
                        sheets=[SheetConfig(name="Sheet1", workbook_sheet="Sheet1")],
                    )
                ],
            )
        )
        app.refresh_tree()
        app.update_idletasks()

        # Source node(s) should be open
        for s_id in app.tree.get_children(""):
            assert app.tree.item(s_id, "open") in (True, 1, "1")
            # Recipe nodes should be open too
            for r_id in app.tree.get_children(s_id):
                assert app.tree.item(r_id, "open") in (True, 1, "1")
    finally:
        app.destroy()
