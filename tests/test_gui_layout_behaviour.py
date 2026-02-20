"""Tests for gui.app UI layout — button order/style, panel visibility, tree expand, move actions."""
import tkinter as tk
from tkinter import ttk

import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


# ---- Helpers ----

def _find_button(root: tk.Misc, text: str) -> ttk.Button | None:
    for w in root.winfo_children():
        if isinstance(w, ttk.Button) and w.cget("text") == text:
            return w
        if isinstance(w, (tk.Frame, ttk.Frame, ttk.LabelFrame)):
            found = _find_button(w, text)
            if found is not None:
                return found
    return None


def _topbar_order(a) -> list[str]:
    btn = _find_button(a, "Add Source (XLSX/CSV)")
    assert btn is not None
    topbar = btn.master
    slaves = sorted(topbar.grid_slaves(row=0), key=lambda w: int(w.grid_info().get("column", 0)))
    labels = []
    for w in slaves:
        if isinstance(w, ttk.Button):
            labels.append(w.cget("text"))
        elif isinstance(w, (tk.Frame, ttk.Frame)):
            kids = sorted(w.grid_slaves(), key=lambda x: int(x.grid_info().get("row", 0)))
            for k in kids:
                if isinstance(k, ttk.Button):
                    labels.append(k.cget("text"))
    return labels


def _select(tree: ttk.Treeview, item_id: str) -> None:
    tree.selection_set(item_id)
    tree.focus(item_id)


def _child(tree: ttk.Treeview, parent: str, text: str) -> str:
    for iid in tree.get_children(parent):
        if tree.item(iid, "text") == text:
            return iid
    raise AssertionError(f"Tree child not found under {parent!r}: {text!r}")


def _invoke(a: tk.Misc, text: str) -> None:
    btn = _find_button(a, text)
    assert btn is not None, f"Button not found: {text}"
    btn.invoke()


# ---- Button layout ----

def test_add_source_and_run_share_accent_style():
    gui = app.TurboExtractorApp()
    try:
        btn_add = _find_button(gui, "Add Source (XLSX/CSV)")
        btn_run = _find_button(gui, "RUN")
        assert btn_add is not None and btn_run is not None
        assert btn_add.cget("style") == btn_run.cget("style")
    finally:
        gui.destroy()


def test_topbar_button_order():
    gui = app.TurboExtractorApp()
    try:
        order = _topbar_order(gui)
        assert order[:6] == [
            "Add Source (XLSX/CSV)",
            "MOVE ▲", "MOVE ▼",
            "Add Recipe", "Add Sheet", "Remove Selected",
        ]
    finally:
        gui.destroy()


def test_run_buttons_order_run_all_left_run_right():
    gui = app.TurboExtractorApp()
    try:
        btn_all = _find_button(gui, "RUN ALL")
        btn_run = _find_button(gui, "RUN")
        assert btn_all is not None and btn_run is not None
        # Both are packed side=left; RUN ALL should have a lower x position
        gui.update_idletasks()
        assert btn_all.winfo_x() < btn_run.winfo_x()
    finally:
        gui.destroy()


# ---- Tree expand ----

def test_tree_fully_expanded_by_default():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[RecipeConfig(name="Recipe1",
                                  sheets=[SheetConfig(name="Sheet1", workbook_sheet="Sheet1")])],
        ))
        gui.refresh_tree()
        gui.update_idletasks()
        for s_id in gui.tree.get_children(""):
            assert gui.tree.item(s_id, "open") in (True, 1, "1")
            for r_id in gui.tree.get_children(s_id):
                assert gui.tree.item(r_id, "open") in (True, 1, "1")
    finally:
        gui.destroy()


# ---- Right-panel visibility ----

def test_right_panel_source_recipe_shows_selection_hides_editor():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[RecipeConfig(name="Recipe1",
                                  sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
        ))
        gui.refresh_tree(); gui.update_idletasks()

        src_id = gui.tree.get_children()[0]
        rec_id = _child(gui.tree, src_id, "Recipe1")
        sh_id  = _child(gui.tree, rec_id, "S1")

        _select(gui.tree, src_id); gui._on_tree_select()
        assert gui.selection_box.winfo_ismapped()
        assert not gui.sheet_box.winfo_ismapped()

        _select(gui.tree, rec_id); gui._on_tree_select()
        assert gui.selection_box.winfo_ismapped()
        assert not gui.sheet_box.winfo_ismapped()

        _select(gui.tree, sh_id); gui._on_tree_select()
        assert gui.sheet_box.winfo_ismapped()
        assert gui.rules_box.winfo_ismapped()
        assert gui.dest_box.winfo_ismapped()
        # selection box should be hidden when sheet is active
        assert not gui.selection_box.winfo_ismapped()
    finally:
        gui.destroy()


# ---- Add recipe / sheet ----

def test_add_recipe_from_source_recipe_or_sheet():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[RecipeConfig(name="Recipe1",
                                  sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
        ))
        gui.refresh_tree(); gui.update_idletasks()

        src_id = gui.tree.get_children()[0]
        rec_id = _child(gui.tree, src_id, "Recipe1")
        sh_id  = _child(gui.tree, rec_id, "S1")

        for node in (src_id, rec_id, sh_id):
            before = len(gui.project.sources[0].recipes)
            _select(gui.tree, node); gui._on_tree_select()
            gui.add_recipe()
            assert len(gui.project.sources[0].recipes) == before + 1
    finally:
        gui.destroy()


def test_add_sheet_default_name_is_Sheet1():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(SourceConfig(
            path="C:/tmp/example.xlsx",
            recipes=[RecipeConfig(name="Recipe1",
                                  sheets=[SheetConfig(name="S1", workbook_sheet="S1")])],
        ))
        gui.refresh_tree(); gui.update_idletasks()

        src_id = gui.tree.get_children()[0]
        rec_id = _child(gui.tree, src_id, "Recipe1")
        sh_id  = _child(gui.tree, rec_id, "S1")

        for node in (src_id, rec_id, sh_id):
            _select(gui.tree, node); gui._on_tree_select()
            gui.add_sheet()
            assert gui.project.sources[0].recipes[0].sheets[-1].name == "Sheet1"
    finally:
        gui.destroy()


# ---- Move recipe and sheet ----

def test_move_recipe_and_sheet_up_down():
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="C:/tmp/example.xlsx", recipes=[
            RecipeConfig(name="Recipe1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1"),
                SheetConfig(name="S2", workbook_sheet="S2"),
            ]),
            RecipeConfig(name="Recipe2", sheets=[
                SheetConfig(name="T1", workbook_sheet="T1"),
                SheetConfig(name="T2", workbook_sheet="T2"),
            ]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree(); gui.update_idletasks()

        # Move Recipe2 up
        src_id = gui.tree.get_children()[0]
        r2_id = _child(gui.tree, src_id, "Recipe2")
        _select(gui.tree, r2_id); gui._on_tree_select(); gui.update_idletasks()
        _invoke(gui, "MOVE ▲"); gui.update_idletasks()
        assert gui.project.sources[0].recipes[0].name == "Recipe2"

        # Move T2 up within Recipe2 (now first)
        src_id = gui.tree.get_children()[0]
        r2_id = _child(gui.tree, src_id, "Recipe2")
        t2_id = _child(gui.tree, r2_id, "T2")
        _select(gui.tree, t2_id); gui._on_tree_select(); gui.update_idletasks()
        _invoke(gui, "MOVE ▲"); gui.update_idletasks()
        assert gui.project.sources[0].recipes[0].sheets[0].name == "T2"

        # Move T2 back down
        src_id = gui.tree.get_children()[0]
        r2_id = _child(gui.tree, src_id, "Recipe2")
        t2_id = _child(gui.tree, r2_id, "T2")
        _select(gui.tree, t2_id); gui._on_tree_select(); gui.update_idletasks()
        _invoke(gui, "MOVE ▼"); gui.update_idletasks()
        assert gui.project.sources[0].recipes[0].sheets[1].name == "T2"
    finally:
        gui.destroy()
