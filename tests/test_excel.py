# Tests for excel.py. Each test checks one behaviour and is named after it.

import openpyxl
import pytest

import excel
import templates
from fakes import PROMPT


# ---------- format_output ----------

def test_format_output_adds_echo_line_by_default():
    block = excel.format_output(PROMPT, "show x", ["out"], None, True, 1)
    assert block == [PROMPT + "show x", "out"]


def test_format_output_skip_first_echo_drops_echo_for_first_port():
    block = excel.format_output(PROMPT, "show x", ["out"], "skip_first_echo", True, 1)
    assert block == ["out"]


def test_format_output_skip_first_echo_keeps_echo_for_later_ports():
    block = excel.format_output(PROMPT, "show x", ["out"], "skip_first_echo", False, 1)
    assert block == [PROMPT + "show x", "out"]


def test_format_output_header_starts_at_port_line_without_echo():
    lines = ["", "note line", "Port   Temp", "Et1/1  35.2"]
    block = excel.format_output(PROMPT, "show int tx", lines, "header", True, 1)
    assert block == ["Port   Temp", "Et1/1  35.2"]


def test_format_output_header_missing_raises_naming_command():
    with pytest.raises(ValueError, match="show int tx"):
        excel.format_output(PROMPT, "show int tx", ["% Invalid input"], "header", True, 1)


def test_format_output_header_with_no_output_raises():
    with pytest.raises(ValueError, match="show int tx"):
        excel.format_output(PROMPT, "show int tx", [], "header", True, 1)


def test_format_output_no_output_gets_one_placeholder_line():
    # A port with no module prints nothing. One blank line keeps the later ports in place.
    block = excel.format_output(PROMPT, "show x", [], None, False, 1)
    assert block == [PROMPT + "show x", " "]


def test_format_output_no_output_on_first_port_without_echo_gets_placeholder():
    block = excel.format_output(PROMPT, "show x", [], "skip_first_echo", True, 1)
    assert block == [" "]


def test_format_output_no_output_gets_one_placeholder_per_expected_line():
    # An 800G SN block is 2 lines, so an empty port still needs 2 rows.
    block = excel.format_output(PROMPT, "show x", [], None, False, 2)
    assert block == [PROMPT + "show x", " ", " "]


def test_format_output_short_output_is_treated_as_an_empty_port():
    # With no module, 800G "grep -i SN" prints only the "Lane SNR" line, 1 of 2. The whole
    # block becomes blank, so the Summary shows a blank SN and not the header text.
    lines = ["    Lane  SNR(dB)    VEYE(U,M,L)"]
    block = excel.format_output(PROMPT, "show x", lines, None, False, 2)
    assert block == [PROMPT + "show x", " ", " "]


def test_format_output_full_output_is_kept_as_is():
    lines = ["  Transceiver SN   SN01", "    Lane  SNR(dB)    VEYE(U,M,L)"]
    block = excel.format_output(PROMPT, "show x", lines, None, False, 2)
    assert block == [PROMPT + "show x"] + lines


def test_format_output_more_lines_than_expected_raises_naming_command():
    # One extra line would shift every later port onto the wrong row, so the run stops.
    lines = ["  Transceiver SN   SN01", "  an extra line"]
    with pytest.raises(ValueError, match="show x"):
        excel.format_output(PROMPT, "show x", lines, None, False, 1)


def test_format_output_any_length_step_keeps_all_lines():
    # lines_per_port None is for tables and histograms, which can be any length.
    block = excel.format_output(PROMPT, "show x", ["a", "b", "c"], None, False, None)
    assert block == [PROMPT + "show x", "a", "b", "c"]


def test_format_output_any_length_step_with_no_output_gets_one_blank_line():
    block = excel.format_output(PROMPT, "show x", [], None, False, None)
    assert block == [PROMPT + "show x", " "]


# ---------- paste ----------

def test_paste_writes_lines_down_one_column():
    worksheet = openpyxl.Workbook().active
    excel.paste(worksheet, "G2", ["a", "b", "c"])
    assert worksheet["G2"].value == "a"
    assert worksheet["G3"].value == "b"
    assert worksheet["G4"].value == "c"
    assert worksheet["G5"].value is None


def test_paste_keeps_line_starting_with_equals_as_text(tmp_path):
    workbook = openpyxl.Workbook()
    excel.paste(workbook.active, "A1", ["=== not a formula ==="])
    workbook.save(tmp_path / "t.xlsx")
    cell = openpyxl.load_workbook(tmp_path / "t.xlsx").active["A1"]
    assert cell.value == "=== not a formula ==="
    assert cell.data_type == "s"


def test_paste_placeholder_survives_save_as_blank_text(tmp_path):
    # A formula that points at an empty cell shows 0 in Excel. A " " cell shows blank.
    workbook = openpyxl.Workbook()
    excel.paste(workbook.active, "A1", [" "])
    workbook.save(tmp_path / "t.xlsx")
    cell = openpyxl.load_workbook(tmp_path / "t.xlsx").active["A1"]
    assert cell.value == " "


# ---------- paste_columns ----------

def test_paste_columns_puts_each_block_down_its_own_column_group():
    worksheet = openpyxl.Workbook().active
    excel.paste_columns(worksheet, "B5", [["a1", "a2"], ["b1"], ["c1"]])
    assert worksheet["B5"].value == "a1"
    assert worksheet["B6"].value == "a2"
    assert worksheet["G5"].value == "b1"
    assert worksheet["L5"].value == "c1"
    assert worksheet["C5"].value is None


# ---------- get_cmds_from_excel (real templates) ----------

def test_get_cmds_200g_step1_reads_32_ports_in_order():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["200G"])
    step = templates.TEMPLATES["200G"][0]
    cmds = excel.get_cmds_from_excel(workbook, step.tab, step.commands)
    assert len(cmds) == 32
    assert cmds[0] == "show interfaces Et1/1 phy detail | grep -i SN"
    assert cmds[31] == "show interfaces Et32/1 phy detail | grep -i SN"


def test_get_cmds_400g_step4_follows_references_to_meta_tab():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["400G"])
    cmds = excel.get_cmds_from_excel(workbook, "Step-4 ", "B3:B34")
    assert len(cmds) == 32
    assert cmds[0] == "show interfaces Et1/1 phy detail | grep -i Pre-FEC"
    for i in range(len(cmds)):
        assert not cmds[i].startswith("=")


def test_get_cmds_reads_commands_from_the_tab_named_in_the_range():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["800G"])
    cmds = excel.get_cmds_from_excel(workbook, "Result-2", "Test command!O3:O34")
    assert len(cmds) == 32
    assert cmds[0] == "show interfaces Et1/1-$ transceiver dom | grep -i case"
    assert cmds[31] == "show interfaces Et62/1-$ transceiver dom | grep -i case"
