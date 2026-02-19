import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination


def _make_source(path: str) -> SourceConfig:
    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="",
        rows_spec="",
        paste_mode="pack",
        rules_combine="AND",
        rules=[],
        destination=Destination(file_path="", sheet_name="Sheet1", start_col="A", start_row=""),
    )
    recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
    return SourceConfig(path=path, recipes=[recipe])


def test_move_source_up_and_down_reorders_project_and_tree():
    gui = app.TurboExtractorApp()

    s1 = _make_source("a.xlsx")
    s2 = _make_source("b.xlsx")
    gui.project = ProjectConfig(sources=[s1, s2])
    gui.refresh_tree()

    # Select second source
    src_ids = gui.tree.get_children("")
    assert len(src_ids) == 2
    gui.tree.selection_set(src_ids[1])
    gui._on_tree_select()

    gui.move_source_up()

    assert [s.path for s in gui.project.sources] == ["b.xlsx", "a.xlsx"]
    src_ids = gui.tree.get_children("")
    assert gui.tree.item(src_ids[0], "text") == "b.xlsx"

    # Move it back down
    gui.tree.selection_set(src_ids[0])
    gui._on_tree_select()

    gui.move_source_down()

    assert [s.path for s in gui.project.sources] == ["a.xlsx", "b.xlsx"]
    src_ids = gui.tree.get_children("")
    assert gui.tree.item(src_ids[1], "text") == "b.xlsx"

    gui.destroy()
