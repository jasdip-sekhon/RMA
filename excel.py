# 800G Result-5: each port's histogram goes in its own group of 5 columns (B, G, L, ...).
COLUMN_GROUP_WIDTH = 5


def _get_cell_text(workbook, sheet, coord):
    value = workbook[sheet][coord].value
    if value.startswith("="):
        ref_sheet, ref_coord = value[1:].split("!")
        ref_sheet = ref_sheet.strip("'")   # 'META RMA CMD' -> META RMA CMD
        value = workbook[ref_sheet][ref_coord].value
    return value.strip()


def get_cmds_from_excel(workbook, tab, cell_range):
    if "!" in cell_range:
        # 800G keeps every command in one tab: "Test command!D3:D34"
        tab, cell_range = cell_range.split("!")
    cmds = []
    rows = workbook[tab][cell_range]
    for i in range(len(rows)):
        cmds.append(_get_cell_text(workbook, tab, rows[i][0].coordinate))
    return cmds


# Need to rebuild output with echo lines (netmiko strips them):
#   7060DX4-32-F#show interfaces Et1/1 phy detail | grep -i SN
#   └── prompt ─┘└──────────── command ──────────────────────┘
def format_output(prompt, command, lines, special, first, lines_per_port):
    # special and lines_per_port: see Step in templates.py
    if special == "header":
        i = 0
        while i < len(lines) and not lines[i].startswith("Port"):
            i += 1
        if i == len(lines):
            raise ValueError("Header not found in output of: " + command)
        block = lines[i:]
    else:
        if len(lines) < lines_per_port:
            # An empty port (no module) prints nothing, or only part of its block (800G
            # "grep -i SN" still prints the "Lane SNR" header). Placeholder lines, as many
            # as one port normally prints, keep every later port on the row its formula
            # reads; " " (not "") so Excel shows a blank rather than 0.
            placeholder = []
            for i in range(lines_per_port):
                placeholder.append(" ")
            lines = placeholder
        if special == "skip_first_echo" and first:
            block = lines
        else:
            block = [prompt + command] + lines
    return block


def paste(worksheet, first_cell, lines):
    # Writes lines down one column from first_cell. The workbook is always a blank
    # template, so there is no old output to clear first.
    row = worksheet[first_cell].row
    col = worksheet[first_cell].column
    for i in range(len(lines)):
        cell = worksheet.cell(row=row + i, column=col)
        cell.value = lines[i]
        if lines[i].startswith("="):
            cell.data_type = "s"   # keep it as text; openpyxl would store it as a formula


def paste_columns(worksheet, first_cell, blocks):
    # One block (port) per column group: block i goes down the column COLUMN_GROUP_WIDTH * i
    # to the right of first_cell, starting on first_cell's row.
    row = worksheet[first_cell].row
    col = worksheet[first_cell].column
    for i in range(len(blocks)):
        cell = worksheet.cell(row=row, column=col + COLUMN_GROUP_WIDTH * i)
        paste(worksheet, cell.coordinate, blocks[i])
