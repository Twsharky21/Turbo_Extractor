\
import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


def test_add_remove_structure_logic():
    gui = app.TurboExtractorApp()

    # manually create source
    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(
            file_path="",
            sheet_name="Sheet1",
            start_col="A",
            start_row="",
        ),
    )
    recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    # remove sheet → should auto-remove recipe
    src_id = gui.tree.get_children()[0]
    recipe_id = gui.tree.get_children(src_id)[0]
    sheet_id = gui.tree.get_children(recipe_id)[0]

    gui.tree.selection_set(sheet_id)
    gui.remove_selected()

    assert len(gui.project.sources[0].recipes) == 0

    gui.destroy()
