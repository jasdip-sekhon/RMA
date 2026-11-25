# Tests for templates.py, read against the real master workbooks.

import openpyxl

import excel
import templates
from fakes import PORTS_800G, port_of


def test_800g_master_tests_the_switch_ports_split_8_ways():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["800G"])
    cmds = excel.get_cmds_from_excel(workbook, "Result-2", "Test command!O3:O34")
    ports = []
    for i in range(len(cmds)):
        ports.append(port_of(cmds[i]))
    assert ports == PORTS_800G


def test_templates_200g_step6_expects_20_lines_per_port():
    # Step-6 formulas read every 21st row: a command line plus exactly 20 lines per port.
    step = templates.TEMPLATES["200G"][5]
    assert step.tab == "Step-6 (Line protocol)"
    assert step.lines_per_port == 20
