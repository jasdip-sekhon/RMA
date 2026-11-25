# Fakes shared by more than one test file. A fake switch is a plain function in a
# SimpleNamespace, used in place of a netmiko connection.

import re
from types import SimpleNamespace

PROMPT = "7060DX4-32-F#"


def fake_prompt():
    return PROMPT


def _fake_disconnect():
    pass


def fake_connection(send_command):
    return SimpleNamespace(send_command=send_command,
                           find_prompt=fake_prompt,
                           disconnect=_fake_disconnect)


def port_of(cmd):
    # Finds the port number in commands such as "sho int et 7/1" or "show int Et7/1-$".
    match = re.search(r"(?i)et\s*(\d+)/1", cmd)
    number = match.group(1)
    return int(number)


# What a fake switch prints for "show interfaces transceiver". EOS prints some notes and
# a two-line header above the "Port" line. The template's paste starts at that line.
TRANSCEIVER_OUTPUT = [
    "If device is externally calibrated, only calibrated values are printed.",
    "                                         Bias      Optical   Optical",
    "          Temp       Voltage   Current   Tx Power  Rx Power",
    "Port      (Celsius)  (Volts)   (mA)      (dBm)     (dBm)",
    "--------- ---------- --------- --------- --------- ---------",
    "Et1/1     48.38      3.31      83.00     1.31      0.47",
]

SHO_INT_LINES = 20     # 200G Step-6 formulas expect "sho int" to print 20 lines.
HISTOGRAM_LINES = 14   # No formulas read 400G Step-7, so any length works.


def _sho_int_sta_output():
    # "sho int sta" prints a header, then one line per port in order, then one other line.
    lines = ["Port       Name   Status       Vlan     Duplex Speed  Type"]
    for port in range(1, 33):
        line = "Et%d/1             connected    100      full   400G   status-%02d"
        lines.append(line % (port, port))
    lines.append("Et33              notconnect   1        full   10G    Not Present")
    return lines


def fake_switch_reply(cmd):
    # Answers every command the 200G and 400G templates send. Each reply names its port.
    if cmd == "show interfaces transceiver":
        reply = "\n".join(TRANSCEIVER_OUTPUT)
    elif cmd == "sho int sta":
        reply = "\n".join(_sho_int_sta_output())
    elif cmd.startswith("clear"):
        reply = ""
    elif cmd.endswith("grep -i SN"):
        reply = "  Transceiver SN   SN%02d" % port_of(cmd)
    elif cmd.endswith("grep -i Interface"):
        reply = "  Interface state  up-%02d" % port_of(cmd)
    elif cmd.endswith("grep -i Pre-FEC"):
        reply = "  Pre-FEC bit error rate  BER-%02d" % port_of(cmd)
    elif cmd.endswith("grep -i case"):
        reply = "   Case temperature   temp-%02d" % port_of(cmd)
    elif cmd.endswith("error-correction histogram"):
        port = port_of(cmd)
        lines = [" Ethernet%d/1" % port]
        for i in range(1, HISTOGRAM_LINES):
            lines.append("  Bin%d  0" % i)
        reply = "\n".join(lines)
    else:
        port = port_of(cmd)   # A "sho int et N/1" command.
        lines = ["Ethernet%d/1 is up, line protocol is up (connected)" % port]
        for i in range(1, SHO_INT_LINES):
            lines.append("  detail line %d" % i)
        reply = "\n".join(lines)
    return reply


# 800G tests ports in pairs, 1, 2, 5, 6 and so on up to 61, 62. Port 18 is skipped. On the
# lab switch it is one 800G port (it may be damaged) instead of 8 x 100G like the others,
# so the template tests port 19 instead.
PORTS_800G = []
for _k in range(16):
    PORTS_800G.append(4 * _k + 1)
    PORTS_800G.append(4 * _k + 2)
PORTS_800G[PORTS_800G.index(18)] = 19
