\
import os
from tempfile import TemporaryDirectory

from core.project import ProjectConfig, SourceConfig, RecipeConfig
from core.models import SheetConfig, Destination, Rule


def make_sample_project(dest_path: str):
    sheet = SheetConfig(
        name="Sheet1",
        workbook_sheet="Sheet1",
        columns_spec="A,C",
        rows_spec="1-1",
        paste_mode="pack",
        rules_combine="AND",
        rules=[Rule(mode="include", column="A", operator="equals", value="alpha")],
        destination=Destination(
            file_path=dest_path,
            sheet_name="Out",
            start_col="B",
            start_row="",
        ),
    )
    recipe = RecipeConfig(name="R1", sheets=[sheet])
    source = SourceConfig(path="dummy.xlsx", recipes=[recipe])
    return ProjectConfig(sources=[source])


def test_project_serialization_roundtrip():
    with TemporaryDirectory() as td:
        json_path = os.path.join(td, "proj.json")
        proj = make_sample_project("out.xlsx")

        proj.save_json(json_path)
        loaded = ProjectConfig.load_json(json_path)

        assert len(loaded.sources) == 1
        assert loaded.sources[0].recipes[0].name == "R1"
        assert loaded.sources[0].recipes[0].sheets[0].columns_spec == "A,C"


def test_project_build_run_items_order():
    proj = make_sample_project("out.xlsx")
    items = proj.build_run_items()

    assert len(items) == 1
    src_path, recipe_name, sheet = items[0]
    assert src_path == "dummy.xlsx"
    assert recipe_name == "R1"
    assert sheet.name == "Sheet1"
