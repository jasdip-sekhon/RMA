# In 800G Result-5, each port's histogram goes in its own group of 5 columns.
COLUMN_GROUP_WIDTH = 5


def _get_cell_text(workbook, sheet, coord):
    value = workbook[sheet][coord].value
    if value.startswith("="):
        ref_sheet, ref_coord = value[1:].split("!")
        ref_sheet = ref_sheet.strip("'")   # Excel puts quotes around sheet names with spaces.
        value = workbook[ref_sheet][ref_coord].value
    return value.strip()


def get_cmds_from_excel(workbook, tab, cell_range):
    if "!" in cell_range:
        # 800G keeps all its commands in the Test command tab.
        tab, cell_range = cell_range.split("!")
    cmds = []
    rows = workbook[tab][cell_range]
    for i in range(len(rows)):
        cmds.append(_get_cell_text(workbook, tab, rows[i][0].coordinate))
    return cmds


# Netmiko removes the line that shows the command. A PuTTY paste has it, so add it back.
#   7060DX4-32-F#show interfaces Et1/1 phy detail | grep -i SN
#   └── prompt ─┘└──────────── command ──────────────────────┘
def format_output(prompt, command, lines, special, first, lines_per_port):
    # See Step in templates.py for what special and lines_per_port mean.
    if special == "header":
        i = 0
        while i < len(lines) and not lines[i].startswith("Port"):
            i += 1
        if i == len(lines):
            raise ValueError("Header not found in output of: " + command)
        block = lines[i:]
    else:
        expected = 1
        if lines_per_port is not None:
            expected = lines_per_port
            if len(lines) > expected:
                # An extra line would shift every later port onto the wrong row.
                raise ValueError("Expected " + str(expected) + " line(s) per port, got "
                                 + str(len(lines)) + " from: " + command)
        if len(lines) < expected:
            # Pad an empty port with spaces, so later ports stay put and Excel shows blank, not 0.
            placeholder = []
            for i in range(expected):
                placeholder.append(" ")
            lines = placeholder
        if special == "skip_first_echo" and first:
            block = lines
        else:
            block = [prompt + command] + lines
    return block


def paste(worksheet, first_cell, lines):
    # The template is blank, so there is nothing to clear first.
    row = worksheet[first_cell].row
    col = worksheet[first_cell].column
    for i in range(len(lines)):
        cell = worksheet.cell(row=row + i, column=col)
        cell.value = lines[i]
        if lines[i].startswith("="):
            cell.data_type = "s"   # Keep it as text, or openpyxl saves it as a formula.


def paste_columns(worksheet, first_cell, blocks):
    row = worksheet[first_cell].row
    col = worksheet[first_cell].column
    for i in range(len(blocks)):
        cell = worksheet.cell(row=row, column=col + COLUMN_GROUP_WIDTH * i)
        paste(worksheet, cell.coordinate, blocks[i])
