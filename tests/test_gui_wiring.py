"""Tests for gui.app — import safety, project wiring, run wiring, autosave, inline rename."""
import json
import os

import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule, SheetResult, RunReport


# ---- Import safety ----

def test_gui_module_import_safe():
    assert hasattr(app, "TurboExtractorApp")
    assert callable(app.main)


def test_gui_project_attribute_on_instance():
    gui = app.TurboExtractorApp()
    assert hasattr(gui, "project")
    gui.destroy()


# ---- Run wiring ----

def test_run_all_calls_engine_with_project_items(monkeypatch):
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        columns_spec="", rows_spec="",
        paste_mode="pack", rules_combine="AND", rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Sheet1", start_col="A", start_row=""),
    )
    gui.project = ProjectConfig(sources=[
        SourceConfig(path="src.xlsx", recipes=[RecipeConfig(name="R1", sheets=[sheet])])
    ])

    called = {}

    def fake_run_all(items):
        called["items"] = list(items)
        return RunReport(ok=True, results=[])

    monkeypatch.setattr(app, "engine_run_all", fake_run_all)
    monkeypatch.setattr(gui, "_show_scrollable_report_dialog", lambda *a, **k: None)

    gui.run_all()
    assert called["items"] == gui.project.build_run_items()
    gui.destroy()


def test_run_selected_sheet_calls_engine_with_current_context(monkeypatch):
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        columns_spec="", rows_spec="",
        paste_mode="pack", rules_combine="AND", rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Sheet1", start_col="A", start_row=""),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    src_id = gui.tree.get_children()[0]
    rec_id = gui.tree.get_children(src_id)[0]
    sh_id  = gui.tree.get_children(rec_id)[0]
    gui.tree.selection_set(sh_id)
    gui._on_tree_select()

    called = {}

    def fake_run_sheet(source_path, sheet_cfg, recipe_name="Recipe"):
        called["source_path"] = source_path
        called["sheet_cfg"] = sheet_cfg
        called["recipe_name"] = recipe_name
        return SheetResult(
            source_path=source_path, recipe_name=recipe_name,
            sheet_name=sheet_cfg.name,
            dest_file=sheet_cfg.destination.file_path,
            dest_sheet=sheet_cfg.destination.sheet_name,
            rows_written=7, message="OK",
        )

    monkeypatch.setattr(app, "engine_run_sheet", fake_run_sheet)
    monkeypatch.setattr(app.messagebox, "showinfo", lambda *a, **k: None)

    gui.run_selected_sheet()
    assert called["source_path"] == "src.xlsx"
    assert called["sheet_cfg"] is sheet
    assert called["recipe_name"] == "R1"
    gui.destroy()


# ---- Autosave ----

def test_gui_autosave_writes_project(tmp_path, monkeypatch):
    autosave_path = tmp_path / "autosave.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(autosave_path))

    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(
            __import__("core.project", fromlist=["SourceConfig"]).SourceConfig(path="C:/x.csv", recipes=[])
        )
        gui._mark_dirty()
        gui._autosave_now()

        assert autosave_path.exists()
        data = json.loads(autosave_path.read_text(encoding="utf-8"))
        assert data["sources"][0]["path"] == "C:/x.csv"
    finally:
        gui.destroy()


def test_gui_autoload_on_start(tmp_path, monkeypatch):
    autosave_path = tmp_path / "autosave.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(autosave_path))

    from core.autosave import save_project_atomic
    proj = ProjectConfig(sources=[SourceConfig(path="C:/a.xlsx", recipes=[])])
    save_project_atomic(proj, str(autosave_path))

    gui = app.TurboExtractorApp()
    try:
        assert len(gui.project.sources) == 1
        assert gui.project.sources[0].path == "C:/a.xlsx"
    finally:
        gui.destroy()


# ---- Inline rename ----

def test_inline_rename_recipe_updates_model():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(name="Sheet1", workbook_sheet="Sheet1",
                        destination=Destination(file_path=""))
    recipe = RecipeConfig(name="OldRecipe", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    gui._apply_recipe_rename([0, 0], "NewRecipe")
    assert gui.project.sources[0].recipes[0].name == "NewRecipe"
    gui.destroy()


def test_inline_rename_sheet_updates_model_and_workbook_sheet():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(name="OldSheet", workbook_sheet="OldSheet",
                        destination=Destination(file_path=""))
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    gui._apply_sheet_rename([0, 0, 0], "NewSheetName")
    assert sheet.name == "NewSheetName"
    assert sheet.workbook_sheet == "NewSheetName"
    gui.destroy()
