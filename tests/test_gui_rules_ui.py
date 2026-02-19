\
import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


def test_add_rule_updates_model():
    gui = app.TurboExtractorApp()

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
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    src = gui.tree.get_children()[0]
    rec = gui.tree.get_children(src)[0]
    sh = gui.tree.get_children(rec)[0]
    gui.tree.selection_set(sh)
    gui._on_tree_select()

    gui.add_rule()
    assert len(sheet.rules) == 1

    gui.destroy()
