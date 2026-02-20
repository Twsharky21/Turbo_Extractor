"""Smoke tests — verify all core and gui modules import without error."""


def test_core_modules_import():
    import core.engine
    import core.planner
    import core.writer
    import core.models
    import core.errors
    import core.io
    import core.parsing
    import core.rules
    import core.transform
    import core.project
    import core.templates
    import core.autosave


def test_gui_module_imports():
    import gui.app
    import gui.ui_build
