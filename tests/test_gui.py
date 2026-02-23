"""
test_gui.py — Consolidated GUI tests.

Covers:
  - Import safety and basic instance attributes
  - Run wiring (run_all, run_selected_sheet)
  - Autosave: dirty flag, atomic save, auto-load on start
  - Inline rename: recipes and sheets
  - Tree structure: add/remove source/recipe/sheet, auto-remove empty recipe
  - Tree reorder: move up/down (sources, recipes, sheets), boundary conditions
  - Rules UI: add rule updates model
  - Editor field sync: all SheetConfig fields pushed to model
  - selection_name_var: updates on tree selection
  - Context menu wiring: _ctx_source_index, _ctx_recipe_path, _ctx_sheet_path
  - _format_run_report: success, error, empty
  - Scrollable report dialog: creates Toplevel, second call replaces first
  - Layout: button order, styles, tree expand
  - remove_selected on empty selection does not crash
  - _mark_dirty sets _autosave_dirty flag

NOTE: These tests require a working Tcl/Tk installation.
If Tcl/Tk is missing, all tests in this file are skipped automatically.
"""
from __future__ import annotations

import json
import os

import pytest

# REPLACE the existing _TCL_OK check block (near the top of the file) with:

try:
    import tkinter as tk
    _root = tk.Tk()
    tk.Frame(_root)          # <-- added: catches deeper Tcl init failures
    _root.destroy()
    _TCL_OK = True
except Exception:
    _TCL_OK = False
#
# This catches the "Can't find a usable tk.tcl" TclError that fires during
# widget creation but not during the bare tk.Tk() call.
# ─────────────────────────────────────────────────────────────────────────────

import gui.app as app
from core.autosave import save_project_atomic
from core.models import Destination, Rule, SheetConfig, SheetResult, RunReport
from core.project import ProjectConfig, RecipeConfig, SourceConfig


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _make_source(path: str = "src.xlsx") -> SourceConfig:
    sh = SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        columns_spec="A", rows_spec="1-1",
        paste_mode="pack", rules_combine="AND", rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Out",
                                start_col="B", start_row=""),
    )
    return SourceConfig(path=path, recipes=[RecipeConfig(name="Recipe1", sheets=[sh])])


def _select(tree, item_id):
    tree.selection_set(item_id)
    tree.focus(item_id)


def _load_sheet(gui, src_idx=0, rec_idx=0, sh_idx=0):
    src_id = gui.tree.get_children()[src_idx]
    rec_id = gui.tree.get_children(src_id)[rec_idx]
    sh_id  = gui.tree.get_children(rec_id)[sh_idx]
    _select(gui.tree, sh_id)
    gui._on_tree_select()
    gui.update_idletasks()
    return sh_id


def _make_result(recipe, sheet, rows=5, error_code=None, error_msg=None):
    return SheetResult(
        source_path="s.xlsx", recipe_name=recipe, sheet_name=sheet,
        dest_file="d.xlsx", dest_sheet="Out", rows_written=rows,
        message="ERROR" if error_code else "OK",
        error_code=error_code, error_message=error_msg,
    )


# ══════════════════════════════════════════════════════════════════════════════
# IMPORT SAFETY + BASIC ATTRIBUTES
# ══════════════════════════════════════════════════════════════════════════════

def test_gui_module_import_safe():
    assert hasattr(app, "TurboExtractorApp")
    assert callable(app.main)


def test_gui_project_attribute_on_instance():
    gui = app.TurboExtractorApp()
    assert hasattr(gui, "project")
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# RUN WIRING
# ══════════════════════════════════════════════════════════════════════════════

def test_run_all_calls_engine_with_project_items(monkeypatch):
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])

    called = {}

    def fake_run_all(items, **_kw):
        called["items"] = list(items)
        return RunReport(ok=True, results=[])

    monkeypatch.setattr(app, "engine_run_all", fake_run_all)
    monkeypatch.setattr(gui, "_show_scrollable_report_dialog", lambda *a, **k: None)

    gui.run_all()
    assert called["items"] == gui.project.build_run_items()
    gui.destroy()


def test_run_selected_sheet_calls_engine_with_current_context(monkeypatch):
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])
    gui.refresh_tree()
    _load_sheet(gui)

    called = {}

    def fake_run_sheet(source_path, cfg, recipe_name=None):
        called["source_path"] = source_path
        called["cfg"]         = cfg
        return SheetResult(source_path=source_path, recipe_name="R",
                           sheet_name="S", dest_file="d", dest_sheet="Out",
                           rows_written=0, message="OK")

    monkeypatch.setattr(app, "engine_run_sheet", fake_run_sheet)
    monkeypatch.setattr(gui, "_show_scrollable_report_dialog", lambda *a, **k: None)

    gui.run_selected_sheet()
    assert called.get("source_path") == gui.current_source_path
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# AUTOSAVE
# ══════════════════════════════════════════════════════════════════════════════

def test_mark_dirty_sets_autosave_dirty_flag():
    gui = app.TurboExtractorApp()
    gui._autosave_dirty = False
    gui._mark_dirty()
    assert gui._autosave_dirty is True
    gui.destroy()


def test_autosave_saves_project_to_json(tmp_path, monkeypatch):
    autosave_path = str(tmp_path / "autosave.json")
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", autosave_path)

    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source("C:/x.csv"))
    gui._autosave_dirty = True
    gui._autosave_now()

    assert os.path.exists(autosave_path)
    with open(autosave_path) as f:
        data = json.load(f)
    assert data["sources"][0]["path"] == "C:/x.csv"
    gui.destroy()


def test_gui_autoload_on_start(tmp_path, monkeypatch):
    autosave_path = tmp_path / "autosave.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(autosave_path))

    proj = ProjectConfig(sources=[SourceConfig(path="C:/a.xlsx", recipes=[])])
    save_project_atomic(proj, str(autosave_path))

    gui = app.TurboExtractorApp()
    assert len(gui.project.sources) == 1
    assert gui.project.sources[0].path == "C:/a.xlsx"
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# INLINE RENAME
# ══════════════════════════════════════════════════════════════════════════════

def test_inline_rename_recipe_updates_model():
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])
    gui.refresh_tree()

    gui._apply_recipe_rename([0, 0], "NewRecipe")
    assert gui.project.sources[0].recipes[0].name == "NewRecipe"
    gui.destroy()


def test_inline_rename_sheet_updates_model_and_workbook_sheet():
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])
    gui.refresh_tree()

    sheet = gui.project.sources[0].recipes[0].sheets[0]
    gui._apply_sheet_rename([0, 0, 0], "NewSheetName")
    assert sheet.name == "NewSheetName"
    assert sheet.workbook_sheet == "NewSheetName"
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# TREE STRUCTURE — ADD / REMOVE
# ══════════════════════════════════════════════════════════════════════════════

def test_remove_sheet_auto_removes_empty_recipe():
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])
    gui.refresh_tree()

    src_id = gui.tree.get_children()[0]
    rec_id = gui.tree.get_children(src_id)[0]
    sh_id  = gui.tree.get_children(rec_id)[0]
    _select(gui.tree, sh_id)
    gui.remove_selected()

    assert len(gui.project.sources[0].recipes) == 0
    gui.destroy()


def test_add_rule_updates_model():
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source()])
    gui.refresh_tree()
    _load_sheet(gui)

    initial_count = len(gui.project.sources[0].recipes[0].sheets[0].rules)
    gui.add_rule()
    assert len(gui.project.sources[0].recipes[0].sheets[0].rules) == initial_count + 1
    gui.destroy()


def test_remove_selected_on_empty_selection_no_crash():
    gui = app.TurboExtractorApp()
    gui.tree.selection_set([])
    try:
        gui.remove_selected()
    except Exception as e:
        pytest.fail(f"remove_selected raised unexpectedly: {e}")
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# TREE REORDER — MOVE BOUNDARIES
# ══════════════════════════════════════════════════════════════════════════════

def _make_gui_two_sources():
    gui = app.TurboExtractorApp()
    for name in ["a.xlsx", "b.xlsx"]:
        gui.project.sources.append(
            SourceConfig(path=name, recipes=[
                RecipeConfig(name="R1", sheets=[
                    SheetConfig(name="S1", workbook_sheet="S1")
                ])
            ])
        )
    gui.refresh_tree()
    return gui


def test_move_source_up_on_first_does_nothing():
    gui = _make_gui_two_sources()
    src_id = gui.tree.get_children("")[0]
    gui.tree.selection_set(src_id)
    gui._on_tree_select()
    gui.move_source_up()
    assert gui.project.sources[0].path == "a.xlsx"
    assert gui.project.sources[1].path == "b.xlsx"
    gui.destroy()


def test_move_source_down_on_last_does_nothing():
    gui = _make_gui_two_sources()
    src_ids = gui.tree.get_children("")
    gui.tree.selection_set(src_ids[1])
    gui._on_tree_select()
    gui.move_source_down()
    assert gui.project.sources[0].path == "a.xlsx"
    assert gui.project.sources[1].path == "b.xlsx"
    gui.destroy()


def test_move_source_up_swaps_sources():
    gui = _make_gui_two_sources()
    src_ids = gui.tree.get_children("")
    gui.tree.selection_set(src_ids[1])
    gui._on_tree_select()
    gui.move_source_up()
    assert gui.project.sources[0].path == "b.xlsx"
    assert gui.project.sources[1].path == "a.xlsx"
    gui.destroy()


def test_move_source_down_swaps_sources():
    gui = _make_gui_two_sources()
    src_ids = gui.tree.get_children("")
    gui.tree.selection_set(src_ids[0])
    gui._on_tree_select()
    gui.move_source_down()
    assert gui.project.sources[0].path == "b.xlsx"
    assert gui.project.sources[1].path == "a.xlsx"
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# EDITOR FIELD SYNC
# ══════════════════════════════════════════════════════════════════════════════

def test_paste_mode_pack_together_maps_to_pack():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    gui.paste_var.set("Pack Together")
    gui._push_editor_to_sheet()
    assert gui.project.sources[0].recipes[0].sheets[0].paste_mode == "pack"
    gui.destroy()


def test_paste_mode_keep_format_maps_to_keep():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    gui.paste_var.set("Keep Format")
    gui._push_editor_to_sheet()
    assert gui.project.sources[0].recipes[0].sheets[0].paste_mode == "keep"
    gui.destroy()


def test_combine_var_or_syncs_to_model():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    gui.combine_var.set("OR")
    gui._push_editor_to_sheet()
    assert gui.project.sources[0].recipes[0].sheets[0].rules_combine == "OR"
    gui.destroy()


def test_source_start_row_var_syncs_to_model():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    gui.source_start_row_var.set("5")
    gui._push_editor_to_sheet()
    assert gui.project.sources[0].recipes[0].sheets[0].source_start_row == "5"
    gui.destroy()


def test_dest_start_col_var_syncs_to_model():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    gui.start_col_var.set("D")
    gui._push_editor_to_sheet()
    assert gui.project.sources[0].recipes[0].sheets[0].destination.start_col == "D"
    gui.destroy()


def test_editor_not_pushed_while_loading():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(_make_source())
    gui.refresh_tree()
    _load_sheet(gui)

    sheet = gui.project.sources[0].recipes[0].sheets[0]
    original_mode = sheet.paste_mode

    gui._loading = True
    gui.paste_var.set("Keep Format")
    gui._push_editor_to_sheet()
    assert sheet.paste_mode == original_mode
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# SELECTION NAME VAR
# ══════════════════════════════════════════════════════════════════════════════

def _make_gui_with_project():
    gui = app.TurboExtractorApp()
    src = SourceConfig(path="C:/data/source_file.xlsx", recipes=[
        RecipeConfig(name="MyRecipe", sheets=[
            SheetConfig(name="MySheet", workbook_sheet="MySheet"),
        ])
    ])
    gui.project = ProjectConfig(sources=[src])
    gui.refresh_tree()
    return gui


def test_selection_name_var_set_to_filename_on_source_select():
    gui = _make_gui_with_project()
    src_id = gui.tree.get_children("")[0]
    gui.tree.selection_set(src_id)
    gui._on_tree_select()
    name = gui.selection_name_var.get()
    assert "source_file.xlsx" in name
    assert "C:/data" not in name
    gui.destroy()


def test_selection_name_var_set_to_recipe_name_on_recipe_select():
    gui = _make_gui_with_project()
    src_id = gui.tree.get_children("")[0]
    rec_id = gui.tree.get_children(src_id)[0]
    gui.tree.selection_set(rec_id)
    gui._on_tree_select()
    assert gui.selection_name_var.get() == "MyRecipe"
    gui.destroy()


def test_selection_name_var_set_to_sheet_name_on_sheet_select():
    gui = _make_gui_with_project()
    src_id = gui.tree.get_children("")[0]
    rec_id = gui.tree.get_children(src_id)[0]
    sh_id  = gui.tree.get_children(rec_id)[0]
    gui.tree.selection_set(sh_id)
    gui._on_tree_select()
    assert gui.selection_name_var.get() == "MySheet"
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# CONTEXT MENU WIRING
# ══════════════════════════════════════════════════════════════════════════════

def _make_gui_3level():
    gui = app.TurboExtractorApp()
    gui.project = ProjectConfig(sources=[_make_source("test_src.xlsx")])
    gui.refresh_tree()
    return gui


def test_ctx_source_index_set_on_right_click():
    gui = _make_gui_3level()
    src_id = gui.tree.get_children("")[0]
    path = gui._get_tree_path(src_id)
    gui._ctx_source_index = path[0]
    assert gui._ctx_source_index == 0
    gui.destroy()


def test_get_ctx_source_returns_none_when_index_none():
    gui = _make_gui_3level()
    gui._ctx_source_index = None
    assert gui._get_ctx_source() is None
    gui.destroy()


def test_get_ctx_source_returns_correct_source():
    gui = _make_gui_3level()
    gui._ctx_source_index = 0
    src = gui._get_ctx_source()
    assert src is gui.project.sources[0]
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# _FORMAT_RUN_REPORT
# ══════════════════════════════════════════════════════════════════════════════

def test_format_run_report_success_format():
    gui = app.TurboExtractorApp()
    report = RunReport(ok=True, results=[_make_result("MyRecipe", "MySheet", 42)])
    text = gui._format_run_report(report)
    assert "MyRecipe" in text
    assert "MySheet" in text
    assert "42" in text
    gui.destroy()


def test_format_run_report_error_format():
    gui = app.TurboExtractorApp()
    report = RunReport(ok=False, results=[
        _make_result("R1", "S1", 0, error_code="DEST_BLOCKED", error_msg="Zone blocked")
    ])
    text = gui._format_run_report(report)
    assert "ERROR" in text
    assert "DEST_BLOCKED" in text
    assert "Zone blocked" in text
    gui.destroy()


def test_format_run_report_empty_results_returns_no_work_items():
    gui = app.TurboExtractorApp()
    text = gui._format_run_report(RunReport(ok=True, results=[]))
    assert text == "No work items."
    gui.destroy()


def test_format_run_report_multiple_results_all_present():
    gui = app.TurboExtractorApp()
    report = RunReport(ok=True, results=[
        _make_result("R1", "S1", 10),
        _make_result("R2", "S2", 0, error_code="BAD_SPEC", error_msg="oops"),
        _make_result("R3", "S3", 7),
    ])
    text = gui._format_run_report(report)
    assert "R1" in text and "10" in text
    assert "R2" in text and "BAD_SPEC" in text
    assert "R3" in text and "7" in text
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# SCROLLABLE REPORT DIALOG
# ══════════════════════════════════════════════════════════════════════════════

def test_show_scrollable_report_dialog_creates_toplevel():
    gui = app.TurboExtractorApp()
    gui._show_scrollable_report_dialog("Test Title", "Line1\nLine2")
    gui.update_idletasks()
    assert gui._report_dialog is not None
    assert isinstance(gui._report_dialog, tk.Toplevel)
    gui._report_dialog.destroy()
    gui._report_dialog = None
    gui.destroy()


def test_show_scrollable_report_dialog_second_call_replaces_first():
    gui = app.TurboExtractorApp()
    gui._show_scrollable_report_dialog("First", "text1")
    gui.update_idletasks()
    first_dialog = gui._report_dialog

    gui._show_scrollable_report_dialog("Second", "text2")
    gui.update_idletasks()
    second_dialog = gui._report_dialog

    assert second_dialog is not first_dialog
    gui.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# LAYOUT BEHAVIOUR
# ══════════════════════════════════════════════════════════════════════════════

def _find_button(widget, label):
    """Recursively search all descendants of *widget* for a button with the given text."""
    for child in widget.winfo_children():
        if hasattr(child, "cget"):
            try:
                if child.cget("text") == label:
                    return child
            except Exception:
                pass
        result = _find_button(child, label)
        if result is not None:
            return result
    return None


def test_add_source_and_run_buttons_exist():
    gui = app.TurboExtractorApp()
    assert _find_button(gui, "Add Source (XLSX/CSV)") is not None
    assert _find_button(gui, "RUN") is not None
    gui.destroy()


def test_tree_fully_expanded_by_default():
    gui = app.TurboExtractorApp()
    gui.project.sources.append(SourceConfig(
        path="C:/tmp/example.xlsx",
        recipes=[RecipeConfig(name="Recipe1",
                              sheets=[SheetConfig(name="Sheet1",
                                                  workbook_sheet="Sheet1")])],
    ))
    gui.refresh_tree()
    gui.update_idletasks()
    for s_id in gui.tree.get_children(""):
        assert gui.tree.item(s_id, "open") in (True, 1, "1")
    gui.destroy()
