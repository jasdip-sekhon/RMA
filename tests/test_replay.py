# Replays recorded output from real switches through the real code.
#
# The sample workbooks in tests/data/ hold what a person pasted from a real switch. These
# tests split each pasted column back into each command's output and replay it through
# run_steps into the blank master. Then they follow every Summary formula to the cell it
# reads, and that cell must match the sample. Summary formulas only copy cells, so
# matching cells mean a matching Summary.

import openpyxl

import main
import templates
from fakes import fake_connection


def _norm(cmd):
    # "grep -i  Interface " and "grep -i Interface" are the same command to the switch.
    return " ".join(cmd.split())


def _harvest_blocks(worksheet, col, first_row, last_row, prompt, first_cmd):
    # Splits a pasted column into each command's output lines. A row that starts with the
    # prompt starts a new block. In the 800G sample the first command line was pasted
    # without the prompt, so that first block is filed under first_cmd.
    outputs = {}
    cmd = None
    lines = []
    for row in range(first_row, last_row + 1):
        value = worksheet[col + str(row)].value
        if value is None:
            value = ""
        if value.startswith(prompt):
            if cmd is not None:
                outputs[_norm(cmd)] = lines
            cmd = value[len(prompt):]
            lines = []
        elif row == first_row:
            cmd = first_cmd
        else:
            lines.append(value)
    outputs[_norm(cmd)] = lines
    return outputs


def _column_lines(worksheet, col, first_row, last_row):
    lines = []
    for row in range(first_row, last_row + 1):
        lines.append(worksheet[col + str(row)].value)
    return lines


def _fake_replay_switch(outputs):
    # Answers each command with its recorded output. Commands with no recording, such as
    # the clears and the tabs no formula reads, print nothing.
    def send_command(cmd):
        reply = ""
        if _norm(cmd) in outputs:
            reply = "\n".join(outputs[_norm(cmd)])
        return reply

    return fake_connection(send_command)


def _resolve(workbook, sheet, coord):
    # Follows a chain of plain references, such as ='Step-1'!N2 then =$G$3, to the cell
    # that holds a value. Returns that cell's sheet and coordinate.
    value = workbook[sheet][coord].value
    hops = 0
    while isinstance(value, str) and value.startswith("=") and hops < 5:
        ref = value[1:].replace("$", "")
        if "!" in ref:
            ref_sheet, coord = ref.split("!")
            sheet = ref_sheet.strip("'")
        else:
            coord = ref
        value = workbook[sheet][coord].value
        hops = hops + 1
    return sheet, coord


def _summary_mismatches(result, sample, columns):
    # Compares every cell that rows 2 to 33 of these Summary columns finally read.
    compared = 0
    mismatches = []
    for i in range(len(columns)):
        for row in range(2, 34):
            sheet, coord = _resolve(sample, "Summary", columns[i] + str(row))
            compared = compared + 1
            if result[sheet][coord].value != sample[sheet][coord].value:
                mismatches.append((sheet, coord, result[sheet][coord].value))
    return compared, mismatches


def test_replay_400g_sample_gives_the_same_summary_inputs():
    sample = openpyxl.load_workbook("tests/data/400G_sample.xlsx")
    prompt = "DCS-7060DX5-32-F#"
    outputs = {}
    outputs.update(_harvest_blocks(sample["Step-1"], "G", 2, 65, prompt, None))
    outputs.update(_harvest_blocks(sample["Step-4 "], "G", 2, 65, prompt, None))
    outputs.update(_harvest_blocks(sample["Step-6"], "D", 2, 65, prompt, None))
    outputs["sho int sta"] = _column_lines(sample["Step-3"], "G", 2, 35)
    outputs["show interfaces transceiver"] = _column_lines(sample["Step-5"], "D", 5, 136)

    result = openpyxl.load_workbook(templates.TEMPLATE_FILES["400G"])
    main.run_steps(_fake_replay_switch(outputs), result, prompt, templates.TEMPLATES["400G"])

    compared, mismatches = _summary_mismatches(result, sample, ["C", "F", "G", "H", "I"])
    assert compared == 160
    assert mismatches == []


def test_replay_800g_sample_gives_the_same_summary_inputs():
    sample = openpyxl.load_workbook("tests/data/800G_sample.xlsx")
    prompt = "localhost#"
    outputs = {}
    outputs.update(_harvest_blocks(sample["Result-2"], "B", 2, 65, prompt,
                                   "show interfaces Et1/1-$ transceiver dom | grep -i case"))
    outputs.update(_harvest_blocks(sample["Result-3"], "B", 2, 97, prompt,
                                   "show interfaces Et1/1 phy detail | grep -i SN"))

    result = openpyxl.load_workbook(templates.TEMPLATE_FILES["800G"])
    main.run_steps(_fake_replay_switch(outputs), result, prompt, templates.TEMPLATES["800G"])

    compared, mismatches = _summary_mismatches(result, sample, ["C", "F", "G"])
    assert compared == 96
    assert mismatches == []
