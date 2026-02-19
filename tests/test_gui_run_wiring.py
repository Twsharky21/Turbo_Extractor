import gui.app as app

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, SheetResult, RunReport


def test_run_all_calls_engine_with_project_items(monkeypatch):
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Sheet1", start_col="A", start_row=""),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])

    called = {}

    def fake_run_all(items):
        called["items"] = list(items)
        return RunReport(ok=True, results=[])

    monkeypatch.setattr(app, "engine_run_all", fake_run_all)
    monkeypatch.setattr(app.messagebox, "showinfo", lambda *a, **k: None)

    gui.run_all()

    assert called["items"] == gui.project.build_run_items()
    gui.destroy()


def test_run_selected_sheet_calls_engine_with_current_context(monkeypatch):
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Sheet1", start_col="A", start_row=""),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    # Select sheet node to establish current context
    src_id = gui.tree.get_children()[0]
    recipe_id = gui.tree.get_children(src_id)[0]
    sheet_id = gui.tree.get_children(recipe_id)[0]
    gui.tree.selection_set(sheet_id)
    gui._on_tree_select()

    called = {}

    def fake_run_sheet(source_path, sheet_cfg, recipe_name="Recipe"):
        called["source_path"] = source_path
        called["sheet_cfg"] = sheet_cfg
        called["recipe_name"] = recipe_name
        return SheetResult(
            source_path=source_path,
            recipe_name=recipe_name,
            sheet_name=sheet_cfg.name,
            dest_file=sheet_cfg.destination.file_path,
            dest_sheet=sheet_cfg.destination.sheet_name,
            rows_written=7,
            message="OK",
        )

    monkeypatch.setattr(app, "engine_run_sheet", fake_run_sheet)
    monkeypatch.setattr(app.messagebox, "showinfo", lambda *a, **k: None)

    gui.run_selected_sheet()

    assert called["source_path"] == "src.xlsx"
    assert called["sheet_cfg"] is sheet
    assert called["recipe_name"] == "R1"
    gui.destroy()
