\
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination
import gui.app as app


def test_sheet_selection_binds_editor_and_updates_model():
    gui = app.TurboExtractorApp()

    # Build simple project
    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="A",
        rows_spec="1-1",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(
            file_path="out.xlsx",
            sheet_name="Out",
            start_col="B",
            start_row="",
        ),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])

    gui.refresh_tree()

    # Select sheet node
    src_id = gui.tree.get_children()[0]
    recipe_id = gui.tree.get_children(src_id)[0]
    sheet_id = gui.tree.get_children(recipe_id)[0]

    gui.tree.selection_set(sheet_id)
    gui._on_tree_select()

    # Modify editor field
    gui.columns_var.set("A,C")
    assert sheet.columns_spec == "A,C"

    gui.destroy()
