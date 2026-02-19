\
def test_gui_module_import_safe():
    """
    Importing gui.app should NOT auto-launch Tkinter.
    We only check that the symbols exist and import succeeds.
    """
    import gui.app as app

    assert hasattr(app, "TurboExtractorApp")
    assert callable(app.main)
