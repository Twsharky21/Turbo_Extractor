import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


def test_inline_rename_recipe_updates_model():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(file_path=""),
    )
    recipe = RecipeConfig(name="OldRecipe", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    # Apply rename directly (commit path logic is UI; this is state edit)
    gui._apply_recipe_rename([0, 0], "NewRecipe")
    assert gui.project.sources[0].recipes[0].name == "NewRecipe"

    gui.destroy()


def test_inline_rename_sheet_updates_model_and_workbook_sheet():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(
        name="OldSheet",
        workbook_sheet="OldSheet",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(file_path=""),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    gui._apply_sheet_rename([0, 0, 0], "NewSheetName")
    assert sheet.name == "NewSheetName"
    assert sheet.workbook_sheet == "NewSheetName"

    gui.destroy()
