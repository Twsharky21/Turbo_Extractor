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

    labels: list[str] = []
    for w in slaves:
        if isinstance(w, ttk.Button):
            labels.append(w.cget("text"))
        elif isinstance(w, (tk.Frame, ttk.Frame)):
            # Move frame: include its children in visual order (top then bottom)
            kids = list(w.grid_slaves())
            # grid_slaves returns reverse; sort by row
            kids.sort(key=lambda x: int(x.grid_info().get("row", 0)))
            for k in kids:
                if isinstance(k, ttk.Button):
                    labels.append(k.cget("text"))

    return labels


def _select_and_focus(tree: ttk.Treeview, item_id: str) -> None:
    """Mimic real user click: selection + focus."""
    tree.selection_set(item_id)
    tree.focus(item_id)
    try:
        tree.see(item_id)
    except Exception:
        # In tests the widget may not be fully realized; ignore.
        pass


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
            "MOVE ▲",
            "MOVE ▼",
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


def _find_tree_child_by_text(tree: ttk.Treeview, parent_id: str, text: str) -> str:
    """Return the child iid whose displayed text matches exactly."""
    for iid in tree.get_children(parent_id):
        if tree.item(iid, "text") == text:
            return iid
    raise AssertionError(f"Tree child not found under {parent_id!r}: {text!r}")


def _invoke_button(app: tk.Misc, text: str) -> None:
    btn = _find_button_by_text(app, text)
    assert btn is not None, f"Button not found: {text}"
    btn.invoke()


def test_move_recipe_and_sheet_up_down_reorders_model():
    """Verify MOVE ▲/▼ works for Recipe and Sheet nodes.

    Important testing detail:
    - MOVE handlers may rely on the same selection+focus behavior as a real click.
    - We therefore set selection + focus + call _on_tree_select(), then invoke the button.
    - Treeview item IDs are not stable across refresh, so we always re-find nodes by text.
    """
    from gui.app import TurboExtractorApp
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app = TurboExtractorApp()
    try:
        # Build model: 1 source with 2 recipes, each with 2 sheets
        src = SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[
                RecipeConfig(
                    name="Recipe1",
                    sheets=[
                        SheetConfig(name="S1", workbook_sheet="S1"),
                        SheetConfig(name="S2", workbook_sheet="S2"),
                    ],
                ),
                RecipeConfig(
                    name="Recipe2",
                    sheets=[
                        SheetConfig(name="T1", workbook_sheet="T1"),
                        SheetConfig(name="T2", workbook_sheet="T2"),
                    ],
                ),
            ],
        )
        app.project.sources.append(src)
        app.refresh_tree()
        app.update_idletasks()

        # ---- Move Recipe2 up (becomes first) ----
        source_id = app.tree.get_children("")[0]
        recipe2_id = _find_tree_child_by_text(app.tree, source_id, "Recipe2")
        _select_and_focus(app.tree, recipe2_id)
        # mimic click event wiring in the GUI
        app._on_tree_select()
        app.update_idletasks()

        _invoke_button(app, "MOVE ▲")
        app.update_idletasks()
        assert app.project.sources[0].recipes[0].name == "Recipe2"

        # ---- Move sheet T2 up within Recipe2 ----
        source_id = app.tree.get_children("")[0]
        recipe2_id = _find_tree_child_by_text(app.tree, source_id, "Recipe2")
        t2_id = _find_tree_child_by_text(app.tree, recipe2_id, "T2")
        _select_and_focus(app.tree, t2_id)
        app._on_tree_select()
        app.update_idletasks()

        _invoke_button(app, "MOVE ▲")
        app.update_idletasks()
        assert app.project.sources[0].recipes[0].sheets[0].name == "T2"

        # ---- Move sheet T2 back down ----
        source_id = app.tree.get_children("")[0]
        recipe2_id = _find_tree_child_by_text(app.tree, source_id, "Recipe2")
        t2_id = _find_tree_child_by_text(app.tree, recipe2_id, "T2")
        _select_and_focus(app.tree, t2_id)
        app._on_tree_select()
        app.update_idletasks()

        _invoke_button(app, "MOVE ▼")
        app.update_idletasks()
        assert app.project.sources[0].recipes[0].sheets[1].name == "T2"

    finally:
        app.destroy()


def _build_sample_project_two_sheets(app):
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app.project.sources.clear()
    app.project.sources.append(
        SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[
                RecipeConfig(
                    name="Recipe1",
                    sheets=[
                        SheetConfig(name="S1", workbook_sheet="S1"),
                        SheetConfig(name="S2", workbook_sheet="S2"),
                    ],
                )
            ],
        )
    )
    app.refresh_tree()
    app.update_idletasks()


def test_add_recipe_works_when_clicking_anywhere_in_source_tree():
    from gui.app import TurboExtractorApp
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app = TurboExtractorApp()
    try:
        app.project.sources.append(
            SourceConfig(
                path="C:/tmp/example.xlsx",
                recipes=[RecipeConfig(name="Recipe1", sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
            )
        )
        app.refresh_tree()
        app.update_idletasks()

        source_id = app.tree.get_children("")[0]
        recipe_id = _find_tree_child_by_text(app.tree, source_id, "Recipe1")
        sheet_id = _find_tree_child_by_text(app.tree, recipe_id, "S1")

        _select_and_focus(app.tree, source_id)
        app._on_tree_select()
        app.add_recipe()
        assert len(app.project.sources[0].recipes) == 2

        _select_and_focus(app.tree, recipe_id)
        app._on_tree_select()
        app.add_recipe()
        assert len(app.project.sources[0].recipes) == 3

        _select_and_focus(app.tree, sheet_id)
        app._on_tree_select()
        app.add_recipe()
        assert len(app.project.sources[0].recipes) == 4
    finally:
        app.destroy()


def test_add_sheet_works_from_source_recipe_or_sheet_and_name_is_sheet1():
    from gui.app import TurboExtractorApp
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app = TurboExtractorApp()
    try:
        app.project.sources.append(
            SourceConfig(
                path="C:/tmp/example.xlsx",
                recipes=[RecipeConfig(name="Recipe1", sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
            )
        )
        app.refresh_tree()
        app.update_idletasks()

        source_id = app.tree.get_children("")[0]
        recipe_id = _find_tree_child_by_text(app.tree, source_id, "Recipe1")
        sheet_id = _find_tree_child_by_text(app.tree, recipe_id, "S1")

        _select_and_focus(app.tree, source_id)
        app._on_tree_select()
        app.add_sheet()
        assert app.project.sources[0].recipes[0].sheets[-1].name == "sheet1"

        _select_and_focus(app.tree, recipe_id)
        app._on_tree_select()
        app.add_sheet()
        assert app.project.sources[0].recipes[0].sheets[-1].name == "sheet1"

        _select_and_focus(app.tree, sheet_id)
        app._on_tree_select()
        app.add_sheet()
        assert app.project.sources[0].recipes[0].sheets[-1].name == "sheet1"
    finally:
        app.destroy()


def test_right_panel_shows_only_name_for_source_or_recipe_selection():
    from gui.app import TurboExtractorApp
    from core.project import SourceConfig, RecipeConfig
    from core.models import SheetConfig

    app = TurboExtractorApp()
    try:
        app.project.sources.append(
            SourceConfig(
                path="C:/tmp/example.xlsx",
                recipes=[RecipeConfig(name="Recipe1", sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
            )
        )
        app.refresh_tree()
        app.update_idletasks()

        source_id = app.tree.get_children("")[0]
        recipe_id = _find_tree_child_by_text(app.tree, source_id, "Recipe1")
        sheet_id = _find_tree_child_by_text(app.tree, recipe_id, "S1")

        _select_and_focus(app.tree, source_id)
        app._on_tree_select()
        assert app.selection_box.winfo_ismapped()
        assert not app.sheet_box.winfo_ismapped()
        assert not app.rules_box.winfo_ismapped()
        assert not app.dest_box.winfo_ismapped()

        _select_and_focus(app.tree, recipe_id)
        app._on_tree_select()
        assert app.selection_box.winfo_ismapped()
        assert not app.sheet_box.winfo_ismapped()
        assert not app.rules_box.winfo_ismapped()
        assert not app.dest_box.winfo_ismapped()

        _select_and_focus(app.tree, sheet_id)
        app._on_tree_select()
        assert app.sheet_box.winfo_ismapped()
        assert app.rules_box.winfo_ismapped()
        assert app.dest_box.winfo_ismapped()
    finally:
        app.destroy()


def test_sheet_editor_values_persist_per_sheet_and_do_not_leak():
    from gui.app import TurboExtractorApp

    app = TurboExtractorApp()
    try:
        _build_sample_project_two_sheets(app)

        source_id = app.tree.get_children("")[0]
        recipe_id = _find_tree_child_by_text(app.tree, source_id, "Recipe1")
        s1_id = _find_tree_child_by_text(app.tree, recipe_id, "S1")
        s2_id = _find_tree_child_by_text(app.tree, recipe_id, "S2")

        _select_and_focus(app.tree, s1_id)
        app._on_tree_select()
        app.columns_var.set("A,C")
        app.rows_var.set("1-3")
        app.dest_sheet_var.set("Out1")
        assert app.project.sources[0].recipes[0].sheets[0].columns_spec == "A,C"

        _select_and_focus(app.tree, s2_id)
        app._on_tree_select()
        app.columns_var.set("B")
        app.rows_var.set("4-5")
        app.dest_sheet_var.set("Out2")
        assert app.project.sources[0].recipes[0].sheets[1].columns_spec == "B"

        _select_and_focus(app.tree, s1_id)
        app._on_tree_select()
        assert app.columns_var.get() == "A,C"
        assert app.rows_var.get() == "1-3"
        assert app.dest_sheet_var.get() == "Out1"

        assert app.project.sources[0].recipes[0].sheets[0].destination.sheet_name == "Out1"
        assert app.project.sources[0].recipes[0].sheets[1].destination.sheet_name == "Out2"
    finally:
        app.destroy()
