import os
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

import excel
import ssh
import templates

SWITCHES = {
    "200G": {"device_type": "arista_eos", "host": "192.168.11.220", "username": "admin", "password": "pxq123"},
    "400G": {"device_type": "arista_eos", "host": "192.168.11.221", "username": "admin", "password": "pxq123"},
    "800G": {"device_type": "arista_eos", "host": "192.168.11.200", "username": "admin", "password": "pxq123"},
    "sonic": {"device_type": "linux", "host": "192.168.11.213", "username": "admin", "password": "password"}
}

# Results are kept outside the code folder, so people can find them without opening it.
RESULTS_FOLDER = Path.home() / "Documents" / "RMA results"


def run_steps(connection, workbook, prompt, steps):
    for i in range(len(steps)):
        step = steps[i]
        cmds = excel.get_cmds_from_excel(workbook, step.tab, step.commands)
        print(step.tab + " " + step.commands + ": " + str(len(cmds)) + " commands")

        blocks = []
        for j in range(len(cmds)):
            lines = ssh.run(connection, cmds[j])
            block = excel.format_output(prompt, cmds[j], lines, step.special, j == 0,
                                        step.lines_per_port)
            blocks.append(block)

        if step.paste is None:
            pass   # The step only runs its commands, such as clearing counters.
        elif step.special == "columns":
            excel.paste_columns(workbook[step.tab], step.paste, blocks)
        else:
            all_lines = []
            for j in range(len(blocks)):
                all_lines = all_lines + blocks[j]
            excel.paste(workbook[step.tab], step.paste, all_lines)


def _parse_args():
    if len(sys.argv) != 2 or sys.argv[1] not in templates.TEMPLATE_FILES:
        raise SystemExit("usage: python main.py 200G|400G|800G")
    return sys.argv[1]


def _run_on_switch(template_key, workbook):
    connection = ssh.create_connection(SWITCHES[template_key])
    try:
        prompt = ssh.get_prompt(connection)
        run_steps(connection, workbook, prompt, templates.TEMPLATES[template_key])
    finally:
        ssh.close(connection)


def _save_copy(template_key, workbook):
    RESULTS_FOLDER.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = RESULTS_FOLDER / (template_key + "_" + timestamp + ".xlsx")
    workbook.save(out)
    print("Saved", out)
    return out


def main():
    template_key = _parse_args()
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES[template_key])
    _run_on_switch(template_key, workbook)
    out = _save_copy(template_key, workbook)
    os.startfile(out)   # Open it in Excel, which must save it once to calculate the Summary.


if __name__ == "__main__":
    main()