\
def test_gui_import_and_project_attribute():
    import gui.app as app

    instance = app.TurboExtractorApp
    assert hasattr(instance, "__init__")

    # Ensure ProjectConfig attribute exists when instantiated (without mainloop)
    gui = instance()
    assert hasattr(gui, "project")
    gui.destroy()
