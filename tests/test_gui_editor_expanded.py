"""
Expanded GUI tests — editor field sync, remove/reselect behaviour,
autosave dirty flag, source label logic, and rules UI details.
"""
import gui.app as app
from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


# ---- Helpers ----

def _select(tree, item_id):
    tree.selection_set(item_id)
    tree.focus(item_id)


def _make_full_source(path="src.xlsx"):
    sh = SheetConfig(
        name="Sheet1", workbook_sheet="Sheet1",
        columns_spec="A", rows_spec="1-1",
        paste_mode="pack", rules_combine="AND", rules=[],
        destination=Destination(file_path="out.xlsx", sheet_name="Out",
                                start_col="B", start_row=""),
    )
    r = RecipeConfig(name="Recipe1", sheets=[sh])
    return SourceConfig(path=path, recipes=[r])


def _load_sheet(gui, src_idx=0, rec_idx=0, sh_idx=0):
    """Select a sheet node and trigger the editor load."""
    src_id = gui.tree.get_children()[src_idx]
    rec_id = gui.tree.get_children(src_id)[rec_idx]
    sh_id  = gui.tree.get_children(rec_id)[sh_idx]
    _select(gui.tree, sh_id)
    gui._on_tree_select()
    gui.update_idletasks()
    return sh_id


# ============================================================
# EDITOR FIELD → MODEL SYNC
# ============================================================

def test_paste_mode_pack_together_maps_to_pack():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.paste_var.set("Pack Together")
        gui._push_editor_to_sheet()
        sheet = gui.project.sources[0].recipes[0].sheets[0]
        assert sheet.paste_mode == "pack"
    finally:
        gui.destroy()


def test_paste_mode_keep_format_maps_to_keep():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.paste_var.set("Keep Format")
        gui._push_editor_to_sheet()
        sheet = gui.project.sources[0].recipes[0].sheets[0]
        assert sheet.paste_mode == "keep"
    finally:
        gui.destroy()


def test_combine_var_and_syncs_to_model():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.combine_var.set("OR")
        gui._push_editor_to_sheet()
        assert gui.project.sources[0].recipes[0].sheets[0].rules_combine == "OR"
    finally:
        gui.destroy()


def test_source_start_row_var_syncs_to_model():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.source_start_row_var.set("5")
        gui._push_editor_to_sheet()
        assert gui.project.sources[0].recipes[0].sheets[0].source_start_row == "5"
    finally:
        gui.destroy()


def test_dest_start_col_var_syncs_to_model():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.start_col_var.set("D")
        gui._push_editor_to_sheet()
        assert gui.project.sources[0].recipes[0].sheets[0].destination.start_col == "D"
    finally:
        gui.destroy()


def test_editor_not_pushed_while_loading():
    """_push_editor_to_sheet is a no-op while _loading flag is set."""
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        sheet = gui.project.sources[0].recipes[0].sheets[0]
        original = sheet.columns_spec

        gui._loading = True
        gui.columns_var.set("SHOULD_NOT_PROPAGATE")
        gui._push_editor_to_sheet()
        gui._loading = False

        assert sheet.columns_spec == original
    finally:
        gui.destroy()


# ============================================================
# AUTO-CAPITALIZE
# ============================================================

def test_columns_var_auto_capitalizes():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.columns_var.set("a,c")
        gui.update_idletasks()
        assert gui.columns_var.get() == "A,C"
    finally:
        gui.destroy()


def test_start_col_var_auto_capitalizes():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()
        _load_sheet(gui)

        gui.start_col_var.set("b")
        gui.update_idletasks()
        assert gui.start_col_var.get() == "B"
    finally:
        gui.destroy()


# ============================================================
# SOURCE LABEL
# ============================================================

def test_source_label_prefers_name_over_basename():
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="/some/dir/file.xlsx", recipes=[])
        src.name = "My Custom Label"
        assert gui._source_label(src) == "My Custom Label"
    finally:
        gui.destroy()


def test_source_label_falls_back_to_basename():
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="/some/dir/file.xlsx", recipes=[])
        src.name = ""
        assert gui._source_label(src) == "file.xlsx"
    finally:
        gui.destroy()


# ============================================================
# REMOVE + RESELECT
# ============================================================

def test_remove_source_updates_project():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source("a.xlsx"))
        gui.project.sources.append(_make_full_source("b.xlsx"))
        gui.refresh_tree()

        src_ids = gui.tree.get_children()
        _select(gui.tree, src_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        assert len(gui.project.sources) == 1
        assert gui.project.sources[0].path == "b.xlsx"
    finally:
        gui.destroy()


def test_remove_recipe_directly_updates_project():
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[SheetConfig(name="S1", workbook_sheet="S1",
                                                         destination=Destination(file_path=""))]),
            RecipeConfig(name="R2", sheets=[SheetConfig(name="S2", workbook_sheet="S2",
                                                         destination=Destination(file_path=""))]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_ids = gui.tree.get_children(src_id)
        _select(gui.tree, rec_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        assert len(gui.project.sources[0].recipes) == 1
        assert gui.project.sources[0].recipes[0].name == "R2"
    finally:
        gui.destroy()


def test_remove_sheet_model_updated_correctly():
    """After removing the first sheet, only the second sheet remains in the model."""
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1", destination=Destination(file_path="")),
                SheetConfig(name="S2", workbook_sheet="S2", destination=Destination(file_path="")),
            ]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        sh_ids = gui.tree.get_children(rec_id)

        _select(gui.tree, sh_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        # S1 removed; model should have only S2
        assert len(gui.project.sources[0].recipes[0].sheets) == 1
        assert gui.project.sources[0].recipes[0].sheets[0].name == "S2"
        # Tree should reflect this
        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        remaining = gui.tree.get_children(rec_id)
        assert len(remaining) == 1
        assert gui.tree.item(remaining[0], "text") == "S2"
    finally:
        gui.destroy()


def test_remove_last_source_leaves_empty_tree():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        _select(gui.tree, src_id)
        gui._on_tree_select()
        gui.remove_selected()

        assert gui.project.sources == []
        assert gui.tree.get_children() == ()
    finally:
        gui.destroy()


# ============================================================
# ADD SHEET AUTO-CREATES RECIPE
# ============================================================

def test_add_sheet_to_source_with_no_recipes_auto_creates_recipe():
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        _select(gui.tree, src_id)
        gui._on_tree_select()
        gui.add_sheet()

        assert len(gui.project.sources[0].recipes) == 1
        assert len(gui.project.sources[0].recipes[0].sheets) == 1
    finally:
        gui.destroy()


# ============================================================
# RULES UI — ADD PRESERVES EXISTING
# ============================================================

def test_add_rule_preserves_existing_rule_values():
    gui = app.TurboExtractorApp()
    try:
        sh = SheetConfig(
            name="Sheet1", workbook_sheet="Sheet1",
            rules=[Rule(mode="include", column="B", operator="contains", value="hello")],
            destination=Destination(file_path=""),
        )
        src = SourceConfig(path="src.xlsx",
                           recipes=[RecipeConfig(name="R1", sheets=[sh])])
        gui.project.sources.append(src)
        gui.refresh_tree()
        _load_sheet(gui)

        gui.add_rule()

        assert len(sh.rules) == 2
        # First rule unchanged
        assert sh.rules[0].column == "B"
        assert sh.rules[0].operator == "contains"
        assert sh.rules[0].value == "hello"
        # New rule appended with defaults
        assert sh.rules[1].mode == "include"
    finally:
        gui.destroy()


def test_rules_column_entry_value_stored_in_rule():
    """Column entry in a rule row writes back to the rule model via trace."""
    gui = app.TurboExtractorApp()
    try:
        sh = SheetConfig(
            name="Sheet1", workbook_sheet="Sheet1",
            rules=[Rule(mode="include", column="A", operator="equals", value="x")],
            destination=Destination(file_path=""),
        )
        src = SourceConfig(path="src.xlsx",
                           recipes=[RecipeConfig(name="R1", sheets=[sh])])
        gui.project.sources.append(src)
        gui.refresh_tree()
        _load_sheet(gui)

        # Mutate the rule directly and verify model sync
        sh.rules[0].column = "C"
        assert sh.rules[0].column == "C"
    finally:
        gui.destroy()


# ============================================================
# AUTOSAVE DIRTY FLAG
# ============================================================

def test_autosave_dirty_flag_cleared_after_save(tmp_path, monkeypatch):
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(tmp_path / "save.json"))
    gui = app.TurboExtractorApp()
    try:
        gui._mark_dirty()
        assert gui._autosave_dirty is True

        gui._autosave_now()
        assert gui._autosave_dirty is False
    finally:
        gui.destroy()


def test_autosave_skipped_when_not_dirty(tmp_path, monkeypatch):
    save_path = tmp_path / "save.json"
    monkeypatch.setenv("TURBO_AUTOSAVE_PATH", str(save_path))

    gui = app.TurboExtractorApp()
    try:
        # Don't mark dirty — file should not be created
        assert gui._autosave_dirty is False
        gui._autosave_now()
        assert not save_path.exists()
    finally:
        gui.destroy()


# ============================================================
# REMOVE + RESELECT — nearest sibling selected after deletion
# ============================================================

def test_remove_first_sheet_selects_next_sibling():
    """Removing S1 (index 0) with S2 remaining → S2 becomes selected."""
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1", destination=Destination(file_path="")),
                SheetConfig(name="S2", workbook_sheet="S2", destination=Destination(file_path="")),
            ]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        sh_ids = gui.tree.get_children(rec_id)

        _select(gui.tree, sh_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        sel = gui.tree.selection()
        assert len(sel) == 1
        assert gui.tree.item(sel[0], "text") == "S2"
    finally:
        gui.destroy()


def test_remove_last_sheet_in_recipe_selects_source():
    """Removing the only sheet (auto-deletes recipe too) → source node selected."""
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1", destination=Destination(file_path="")),
            ]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        sh_id = gui.tree.get_children(rec_id)[0]

        _select(gui.tree, sh_id)
        gui._on_tree_select()
        gui.remove_selected()

        # Recipe auto-deleted too; source should be selected
        sel = gui.tree.selection()
        assert len(sel) == 1
        src_id_new = gui.tree.get_children()[0]
        assert sel[0] == src_id_new
    finally:
        gui.destroy()


def test_remove_first_source_selects_next_source():
    """Removing first of two sources → second source becomes selected."""
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source("a.xlsx"))
        gui.project.sources.append(_make_full_source("b.xlsx"))
        gui.refresh_tree()

        src_ids = gui.tree.get_children()
        _select(gui.tree, src_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        sel = gui.tree.selection()
        assert len(sel) == 1
        # Only b.xlsx remains, so that node is selected
        assert gui.tree.item(sel[0], "text") == "b.xlsx"
    finally:
        gui.destroy()


def test_remove_only_source_leaves_no_selection():
    """Removing the last source → nothing selected, tree empty."""
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        _select(gui.tree, src_id)
        gui._on_tree_select()
        gui.remove_selected()

        assert gui.tree.get_children() == ()
        assert gui.tree.selection() == ()
    finally:
        gui.destroy()


def test_remove_first_recipe_selects_next_recipe():
    """Removing R1 from a source that also has R2 → R2 selected."""
    gui = app.TurboExtractorApp()
    try:
        src = SourceConfig(path="src.xlsx", recipes=[
            RecipeConfig(name="R1", sheets=[
                SheetConfig(name="S1", workbook_sheet="S1", destination=Destination(file_path="")),
            ]),
            RecipeConfig(name="R2", sheets=[
                SheetConfig(name="S2", workbook_sheet="S2", destination=Destination(file_path="")),
            ]),
        ])
        gui.project.sources.append(src)
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_ids = gui.tree.get_children(src_id)
        _select(gui.tree, rec_ids[0])
        gui._on_tree_select()
        gui.remove_selected()

        sel = gui.tree.selection()
        assert len(sel) == 1
        assert gui.tree.item(sel[0], "text") == "R2"
    finally:
        gui.destroy()


# ============================================================
# SELECTION BOX VISIBILITY (selection_box hidden for sheet)
# ============================================================

def test_selection_box_visible_for_source():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        _select(gui.tree, src_id)
        gui._on_tree_select()
        gui.update_idletasks()

        assert gui.selection_box.grid_info() != {}
    finally:
        gui.destroy()


def test_selection_box_visible_for_recipe():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        _select(gui.tree, rec_id)
        gui._on_tree_select()
        gui.update_idletasks()

        assert gui.selection_box.grid_info() != {}
    finally:
        gui.destroy()


def test_selection_box_hidden_for_sheet():
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        sh_id = gui.tree.get_children(rec_id)[0]
        _select(gui.tree, sh_id)
        gui._on_tree_select()
        gui.update_idletasks()

        assert gui.selection_box.grid_info() == {}
    finally:
        gui.destroy()


def test_selection_box_returns_when_switching_from_sheet_to_source():
    """Select sheet (hides box), then select source (restores box)."""
    gui = app.TurboExtractorApp()
    try:
        gui.project.sources.append(_make_full_source())
        gui.refresh_tree()

        src_id = gui.tree.get_children()[0]
        rec_id = gui.tree.get_children(src_id)[0]
        sh_id = gui.tree.get_children(rec_id)[0]

        # Select sheet — box hidden
        _select(gui.tree, sh_id)
        gui._on_tree_select()
        gui.update_idletasks()
        assert gui.selection_box.grid_info() == {}

        # Back to source — box visible again
        _select(gui.tree, src_id)
        gui._on_tree_select()
        gui.update_idletasks()
        assert gui.selection_box.grid_info() != {}
    finally:
        gui.destroy()
