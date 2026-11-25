# This file lists each Excel template's file and the steps it runs.

from dataclasses import dataclass
from typing import Optional

# The blank master file for each data rate. The paths are relative to the RMA folder.
TEMPLATE_FILES = {
    "200G": "templates/TH4 Arista command for 200G 2-FR4 testing in Arista TH4_TEMPLATE -.xlsx",
    "400G": "templates/DX5 Arista command for 400G FR4 testing in Arista DX5_TEMPLATE.xlsx",
    "800G": "templates/TH5 Arista command for 800G testing in Arista TH5_TEMPLATE.xlsx",
}


# One step of a template. Data only, and frozen so a row can't be changed by mistake.
@dataclass(frozen=True)
class Step:
    tab: str                            # Where the output goes
    commands: str                       # Such as "B3:B34", or "Test command!D3:D34"
    paste: Optional[str]                # Top cell of the output, or None to paste nothing
    special: Optional[str] = None       # "skip_first_echo", "header" or "columns"
    lines_per_port: Optional[int] = 1   # Exact lines per port, or None for any length


# The steps for each template, in the order they run. The note on each row names the
# formula cells that read its output. "N2 = G3 .. G65" means N2 reads G3, N3 reads G5,
# and so on down to G65.
TEMPLATES = {
    "200G": [
        #    tab                       commands  paste special            lines
        Step("Step-1 (SN)",            "B3:B34", "G2", None),                   # N2 = G3 .. G65
        Step("Step-2 (Clear status)",  "B3:B34", None, None),                   # Clears counters
        Step("Step-3 (Link Status)",   "B3:B34", "G2", None),                   # O2 = G3 .. G65
        Step("Step-4 (Pre-FEC BER)",   "B3:B34", "G2", None),                   # O2 = G3 .. G65
        Step("Step-5 (DOM)",           "A5:A5",  "D5", "header"),               # No formulas
        Step("Step-6 (Line protocol)", "A2:A33", "D2", "skip_first_echo", 20),  # P2 = D2 .. D653
    ],
    "400G": [
        #    tab        commands  paste special            lines
        Step("Step-1",  "B3:B34", "G2", None),                     # N2 = G3 .. G65
        Step("Step-2",  "B3:B34", None, None),                     # Clears counters
        Step("Step-3",  "B2:B2",  "G2", "header"),                 # O2 = G3 .. G34, one table
        Step("Step-4 ", "B3:B34", "G2", None),                     # O2 = G3 .. G65
        Step("Step-5",  "A5:A5",  "D5", "header"),                 # No formulas
        Step("Step-6",  "A3:A34", "D2", None),                     # P2 = D3 .. D65
        Step("Step-7 ", "A3:A34", "D2", "skip_first_echo", None),  # No formulas, any length
    ],
    "800G": [
        # All 800G commands are in the Test command tab. Each "EtN/1-$" command covers
        # 8 lanes. Only Result-2 and Result-3 feed the Summary. Summary F reads Result-2
        # B3, B5 .. B65. Summary C reads Result-5 row 2, which reads Result-3 B3, B6 .. B96.
        #    tab             commands               paste special    lines_per_port
        Step("Test command", "B3:B34",              None, None),           # Clears counters
        Step("Result-1",     "Test command!D3:D34", "B2", None,      8),   # Pre-FEC BER
        Step("Result-1",     "Test command!H3:H34", "F2", None,      8),   # Link status
        Step("Result-1",     "Test command!M3:M34", "K2", None,      8),   # SNR
        Step("Result-2",     "Test command!O3:O34", "B2", None),           # Case temperature
        Step("Result-3",     "Test command!Q3:Q34", "B2", None,      2),   # Serial number
        Step("Result-4",     "Test command!U3:U34", "B2", None,      16),  # TX/RX power
        Step("Result-5",     "Test command!X3:X34", "B5", "columns", None),  # FEC histograms
    ],
}
