def add_sheet(self) -> None:
    sel = self.tree.selection()
    if not sel:
        messagebox.showwarning("Select Recipe", "Select a Recipe to add a Sheet.")
        return

    path = self._get_tree_path(sel[0])
    if len(path) not in (1, 2, 3):
        messagebox.showwarning("Select Recipe", "Select a Source/Recipe/Sheet to add a Sheet.")
        return

    source = self.project.sources[path[0]]

    # Source selected: add under first recipe (create if missing)
    auto_created_recipe = False
    if len(path) == 1:
        if not source.recipes:
            source.recipes.append(RecipeConfig(name="Recipe1", sheets=[]))
            auto_created_recipe = True
        recipe = source.recipes[0]
    else:
        # Recipe selected or Sheet selected -> parent recipe
        recipe = source.recipes[path[1]]

    # Name contract: all new sheets are "sheet1" and duplicates are allowed.
    new_sheet = self._make_default_sheet(name="sheet1")
    recipe.sheets.append(new_sheet)

    if auto_created_recipe:
        # Recipe node didn't exist in tree yet — full refresh required.
        # Incremental insert fails here: r_children would be an empty tuple.
        self.refresh_tree()
    else:
        # Insert into tree incrementally so pre-captured item IDs remain valid.
        src_children = self.tree.get_children("")
        s_id = src_children[path[0]]
        recipe_idx = path[1] if len(path) >= 2 else 0
        r_children = self.tree.get_children(s_id)
        r_id = r_children[recipe_idx]
        self.tree.insert(r_id, "end", text=new_sheet.name)
        self.tree.item(r_id, open=True)
    self._mark_dirty()
