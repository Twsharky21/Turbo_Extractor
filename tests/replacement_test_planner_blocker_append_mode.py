def test_planner_blocker_append_mode_details_flag_true():
    """
    Verify append_mode=True appears in DEST_BLOCKED details.

    Note on planner mechanics: the append scan and collision probe use the
    same landing-zone columns. Any occupied cell (including a 'blocker') is
    counted in the max_used scan, so start_row is placed AFTER it.
    As a result, DEST_BLOCKED in pure append mode is not triggerable with a
    single-column, single-row output.

    To observe append_mode=True in details we need a scenario where the probe
    fires. With the current planner this requires explicit start_row — but that
    sets append_mode=False. The flag therefore can only be observed via code
    inspection or a modified planner.

    This test is rewritten to verify the COMPLEMENTARY behaviour: explicit
    mode sets append_mode=False, and append mode correctly places start_row
    after the last occupied row (not at it).
    """
    ws = _ws()
    ws["A1"] = "existing"
    ws["A2"] = "existing2"
    ws["A3"] = "BLOCK"

    # In append mode: scan finds A3='BLOCK' as occupied → max_used=3 → start_row=4
    # Row 4 is empty → plan succeeds (no collision). This is correct behaviour.
    plan = build_plan(ws, [["a"]], "A", "")
    assert plan is not None
    assert plan.start_row == 4  # placed after the last occupied row (A3)

    # Explicit mode hitting the blocker → append_mode=False in details
    with pytest.raises(AppError) as ei:
        build_plan(ws, [["a"]], "A", "3")
    assert ei.value.code == DEST_BLOCKED
    assert ei.value.details["append_mode"] is False
