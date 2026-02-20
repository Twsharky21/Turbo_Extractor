"""Tests for gui.app tree structure — add/remove, reorder sources, rules UI, sheet editor binding."""
import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


# ---- Helpers ----

def _make_source(path: str) -> SourceConfig:
    sheet = SheetConfig(name="Sheet1", workbook_sheet="Sheet1",
                        destination=Destination(file_path="", sheet_name="Sheet1", start_col="A", start_row=""))
    recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
    return SourceConfig(path=path, recipes=[recipe])


def _select(tree, item_id):
    tree.selection_set(item_id)
    tree.focus(item_id)


# ---- Add / Remove structure ----

def test_remove_sheet_auto_removes_empty_recipe():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(name="Sheet1", workbook_sheet="Sheet1",
                        destination=Destination(file_path=""))
    recipe = RecipeConfig(name="Recipe1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    src_id  = gui.tree.get_children()[0]
    rec_id  = gui.tree.get_children(src_id)[0]
    sh_id   = gui.tree.get_children(rec_id)[0]
    _select(gui.tree, sh_id)
    gui.remove_selected()

    assert len(gui.project.sources[0].recipes) == 0
    gui.destroy()


def test_add_rule_updates_model():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(name="Sheet1", workbook_sheet="Sheet1",
                        columns_spec="", rows_spec="", paste_mode="pack",
                        rules_combine="AND", rules=[],
                        destination=Destination(file_path="", sheet_name="Sheet1", start_col="A", start_row=""))
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    src_id = gui.tree.get_children()[0]
    rec_id = gui.tree.get_children(src_id)[0]
    sh_id  = gui.tree.get_children(rec_id)[0]
    _select(gui.tree, sh_id)
    gui._on_tree_select()

    gui.add_rule()
    assert len(sheet.rules) == 1
    gui.destroy()


# ---- Source reorder ----

def test_move_source_up_and_down_reorders_project_and_tree():
    gui = app.TurboExtractorApp()

    s1 = _make_source("a.xlsx")
    s2 = _make_source("b.xlsx")
    gui.project = ProjectConfig(sources=[s1, s2])
    gui.refresh_tree()

    src_ids = gui.tree.get_children("")
    _select(gui.tree, src_ids[1])
    gui._on_tree_select()
    gui.move_source_up()

    assert [s.path for s in gui.project.sources] == ["b.xlsx", "a.xlsx"]
    src_ids = gui.tree.get_children("")
    assert gui.tree.item(src_ids[0], "text") == "b.xlsx"

    _select(gui.tree, src_ids[0])
    gui._on_tree_select()
    gui.move_source_down()

    assert [s.path for s in gui.project.sources] == ["a.xlsx", "b.xlsx"]
    src_ids = gui.tree.get_children("")
    assert gui.tree.item(src_ids[1], "text") == "b.xlsx"
    gui.destroy()


# ---- Sheet editor binding ----

def test_sheet_selection_binds_editor_and_updates_model():
    gui = app.TurboExtractorApp()

    sheet = SheetConfig(name="Sheet1", workbook_sheet="Sheet1",
                        columns_spec="A", rows_spec="1-1",
                        paste_mode="pack", rules_combine="AND", rules=[],
                        destination=Destination(file_path="out.xlsx", sheet_name="Out", start_col="B", start_row=""))
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="src.xlsx", recipes=[recipe])
    gui.project = ProjectConfig(sources=[source])
    gui.refresh_tree()

    src_id = gui.tree.get_children()[0]
    rec_id = gui.tree.get_children(src_id)[0]
    sh_id  = gui.tree.get_children(rec_id)[0]
    _select(gui.tree, sh_id)
    gui._on_tree_select()

    gui.columns_var.set("A,C")
    assert sheet.columns_spec == "A,C"
    gui.destroy()


def test_sheet_editor_values_do_not_leak_between_sheets():
    gui = app.TurboExtractorApp()

    gui.project.sources.append(SourceConfig(
        path="C:/tmp/example.xlsx",
        recipes=[RecipeConfig(name="Recipe1", sheets=[
            SheetConfig(name="S1", workbook_sheet="S1",
                        destination=Destination(file_path="")),
            SheetConfig(name="S2", workbook_sheet="S2",
                        destination=Destination(file_path="")),
        ])]
    ))
    gui.refresh_tree()
    gui.update_idletasks()

    src_id = gui.tree.get_children()[0]
    rec_id = gui.tree.get_children(src_id)[0]
    s1_id  = gui.tree.get_children(rec_id)[0]
    s2_id  = gui.tree.get_children(rec_id)[1]

    _select(gui.tree, s1_id); gui._on_tree_select()
    gui.columns_var.set("A,C"); gui.rows_var.set("1-3"); gui.dest_sheet_var.set("Out1")

    _select(gui.tree, s2_id); gui._on_tree_select()
    gui.columns_var.set("B"); gui.rows_var.set("4-5"); gui.dest_sheet_var.set("Out2")

    _select(gui.tree, s1_id); gui._on_tree_select()
    assert gui.columns_var.get() == "A,C"
    assert gui.rows_var.get() == "1-3"
    assert gui.dest_sheet_var.get() == "Out1"
    assert gui.project.sources[0].recipes[0].sheets[0].destination.sheet_name == "Out1"
    assert gui.project.sources[0].recipes[0].sheets[1].destination.sheet_name == "Out2"
    gui.destroy()
