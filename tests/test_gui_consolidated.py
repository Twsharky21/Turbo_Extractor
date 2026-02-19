# Consolidated from: test_gui_importsafe.py, test_gui_project_wiring.py, test_gui_rules_ui.py, test_gui_sheet_editor.py, test_gui_tree_structure.py
# Generated: 2026-02-19 20:40 UTC
# NOTE: Function renames applied only to avoid name collisions across original test modules.



# ---- BEGIN test_gui_importsafe.py ----

\
def test_gui_module_import_safe():
    """
    Importing gui.app should NOT auto-launch Tkinter.
    We only check that the symbols exist and import succeeds.
    """
    import gui.app as app

    assert hasattr(app, "TurboExtractorApp")
    assert callable(app.main)


# ---- END {f} ----



# ---- BEGIN test_gui_project_wiring.py ----

\
def test_gui_import_and_project_attribute():
    import gui.app as app

    instance = app.TurboExtractorApp
    assert hasattr(instance, "__init__")

    # Ensure ProjectConfig attribute exists when instantiated (without mainloop)
    gui = instance()
    assert hasattr(gui, "project")
    gui.destroy()


# ---- END {f} ----



# ---- BEGIN test_gui_rules_ui.py ----

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


# ---- END {f} ----



# ---- BEGIN test_gui_sheet_editor.py ----

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


# ---- END {f} ----



# ---- BEGIN test_gui_tree_structure.py ----

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


# ---- END {f} ----
