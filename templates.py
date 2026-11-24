# Fixed facts about each Excel template: which file it is, and its steps.

from collections import namedtuple

# The blank master for each data rate (paths relative to the RMA folder).
TEMPLATE_FILES = {
    "200G": "templates/TH4 Arista command for 200G 2-FR4 testing in Arista TH4_TEMPLATE -.xlsx",
    "400G": "templates/DX5 Arista command for 400G FR4 testing in Arista DX5_TEMPLATE.xlsx",
    "800G": "templates/TH5 Arista command for 800G testing in Arista TH5_TEMPLATE.xlsx",
}

# One step, like a C struct:
#
#   tab            - tab the output is pasted into (and the commands are read from)
#   commands       - cell range holding the commands, e.g. "B3:B34"; may name another
#                    tab, e.g. "Test command!D3:D34"
#   paste          - top cell of the output, e.g. "G2" (None = run only, paste nothing)
#   special        - None, or one of:
#                      "skip_first_echo" - no echo line for the first port
#                      "header"          - no echo line; output starts at the "Port ..." line
#                      "columns"         - one port per column group, 5 columns apart
#   lines_per_port - output lines one port prints (default 1); an empty port gets this
#                    many placeholder lines so later ports stay aligned
Step = namedtuple("Step", ["tab", "commands", "paste", "special", "lines_per_port"],
                  defaults=[1])

# Steps for each template, in run order. Comments name the formula cells that read them.
TEMPLATES = {
    "200G": [
        #    tab                       commands  paste  special
        Step("Step-1 (SN)",            "B3:B34", "G2",  None),               # N2 = G3 .. G65
        Step("Step-2 (Clear status)",  "B3:B34", None,  None),               # run only
        Step("Step-3 (Link Status)",   "B3:B34", "G2",  None),               # O2 = G3 .. G65
        Step("Step-4 (Pre-FEC BER)",   "B3:B34", "G2",  None),               # O2 = G3 .. G65
        Step("Step-5 (DOM)",           "A5:A5",  "D5",  "header"),           # people only
        Step("Step-6 (Line protocol)", "A2:A33", "D2",  "skip_first_echo"),  # P2 = D2 .. D653
    ],
    "400G": [
        #    tab        commands  paste  special
        Step("Step-1",  "B3:B34", "G2",  None),               # N2 = G3 .. G65
        Step("Step-2",  "B3:B34", None,  None),               # run only
        Step("Step-3",  "B2:B2",  "G2",  "header"),           # one table: O2 = G3 .. G34
        Step("Step-4 ", "B3:B34", "G2",  None),               # O2 = G3 .. G65
        Step("Step-5",  "A5:A5",  "D5",  "header"),           # people only
        Step("Step-6",  "A3:A34", "D2",  None),               # P2 = D3 .. D65
        Step("Step-7 ", "A3:A34", "D2",  "skip_first_echo"),  # people only
    ],
    "800G": [
        # All commands live in "Test command"; each "EtN/1-$" command covers 8 lanes.
        #    tab             commands               paste special    lines_per_port
        Step("Test command", "B3:B34",              None, None),             # clear, run only
        Step("Result-1",     "Test command!D3:D34", "B2", None,      8),     # Pre-FEC, people
        Step("Result-1",     "Test command!H3:H34", "F2", None,      8),     # link, people
        Step("Result-1",     "Test command!M3:M34", "K2", None,      8),     # SNR, people
        Step("Result-2",     "Test command!O3:O34", "B2", None),             # Summary F: B3..B65
        Step("Result-3",     "Test command!Q3:Q34", "B2", None,      2),     # R-5 row 2: B3..B96
        Step("Result-4",     "Test command!U3:U34", "B2", None,      16),    # TX/RX, people
        Step("Result-5",     "Test command!X3:X34", "B5", "columns"),        # histogram, people
    ],
}
