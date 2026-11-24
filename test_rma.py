# Tests for the RMA tool. They need no switch. Run them from the RMA folder with pytest.
#
# Each test checks one behaviour and is named after it. Each sets up, acts, then checks.
# The fake switches are plain functions in a SimpleNamespace, used in place of a
# netmiko connection.

import re
from types import SimpleNamespace

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from netmiko import ReadTimeout   # Tests may import netmiko, to fake its errors.

import excel
import main
import ssh
import templates

PROMPT = "7060DX4-32-F#"


def _fake_connection(send_command):
    return SimpleNamespace(send_command=send_command,
                           find_prompt=_fake_prompt,
                           disconnect=_fake_disconnect)


def _fake_prompt():
    return PROMPT


def _fake_disconnect():
    pass


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


def _port_of(cmd):
    # Finds the port number in commands such as "sho int et 7/1" or "show int Et7/1-$".
    match = re.search(r"(?i)et\s*(\d+)/1", cmd)
    number = match.group(1)
    return int(number)


def _sho_int_sta_output():
    # "sho int sta" prints a header, then one line per port in order, then one other line.
    lines = ["Port       Name   Status       Vlan     Duplex Speed  Type"]
    for port in range(1, 33):
        line = "Et%d/1             connected    100      full   400G   status-%02d"
        lines.append(line % (port, port))
    lines.append("Et33              notconnect   1        full   10G    Not Present")
    return lines


def _fake_switch_reply(cmd):
    # Answers every command the 200G and 400G templates send. Each reply names its port.
    if cmd == "show interfaces transceiver":
        reply = "\n".join(TRANSCEIVER_OUTPUT)
    elif cmd == "sho int sta":
        reply = "\n".join(_sho_int_sta_output())
    elif cmd.startswith("clear"):
        reply = ""
    elif cmd.endswith("grep -i SN"):
        reply = "  Transceiver SN   SN%02d" % _port_of(cmd)
    elif cmd.endswith("grep -i Interface"):
        reply = "  Interface state  up-%02d" % _port_of(cmd)
    elif cmd.endswith("grep -i Pre-FEC"):
        reply = "  Pre-FEC bit error rate  BER-%02d" % _port_of(cmd)
    elif cmd.endswith("grep -i case"):
        reply = "   Case temperature   temp-%02d" % _port_of(cmd)
    elif cmd.endswith("error-correction histogram"):
        port = _port_of(cmd)
        lines = [" Ethernet%d/1" % port]
        for i in range(1, HISTOGRAM_LINES):
            lines.append("  Bin%d  0" % i)
        reply = "\n".join(lines)
    else:
        port = _port_of(cmd)   # A "sho int et N/1" command.
        lines = ["Ethernet%d/1 is up, line protocol is up (connected)" % port]
        for i in range(1, SHO_INT_LINES):
            lines.append("  detail line %d" % i)
        reply = "\n".join(lines)
    return reply


def _fake_reply_port5_empty(cmd):
    # Like _fake_switch_reply, but port 5 has no module, so its grep commands print nothing.
    reply = _fake_switch_reply(cmd)
    if "grep" in cmd and _port_of(cmd) == 5:
        reply = ""
    return reply


# 800G tests ports in pairs, 1, 2, 5, 6 and so on up to 61, 62. Port 18 is skipped. On the
# lab switch it is one 800G port (it may be damaged) instead of 8 x 100G like the others,
# so the template tests port 19 instead.
PORTS_800G = []
for _k in range(16):
    PORTS_800G.append(4 * _k + 1)
    PORTS_800G.append(4 * _k + 2)
PORTS_800G[PORTS_800G.index(18)] = 19


def _lane_lines(text, port, count):
    # One line per lane, ending in the port and lane, such as "-05-L3".
    lines = []
    for lane in range(1, count + 1):
        lines.append("%s-%02d-L%d" % (text, port, lane))
    return "\n".join(lines)


def _fake_800g_reply(cmd):
    # Answers every command the 800G template sends. Each "EtN/1-$" command covers 8 lanes.
    port = 0
    if not cmd.startswith("clear"):
        port = _port_of(cmd)
    if cmd.startswith("clear"):
        reply = ""
    elif cmd.endswith("grep -i Pre-FEC"):
        reply = _lane_lines("  Pre-FEC bit error rate  BER", port, 8)
    elif re.search(r"grep -i\s+Interface$", cmd):   # The template uses one or two spaces here.
        reply = _lane_lines("  Interface state  up", port, 8)
    elif cmd.endswith("grep -i SNR"):
        reply = _lane_lines("   SNR  dB", port, 8)
    elif cmd.endswith("grep -i case"):
        reply = "   Case temperature   temp-%02d" % port
    elif cmd.endswith("grep -i SN"):
        # "grep -i SN" also matches the "Lane SNR(dB)" header, so it prints 2 lines.
        reply = "  Transceiver SN   SN%02d\n    Lane  SNR(dB)    VEYE(U,M,L)" % port
    elif cmd.endswith("grep -i optical"):
        reply = _lane_lines("   Optical TX/RX power  dBm", port, 16)
    else:
        lines = [" Ethernet%d/1" % port]
        for i in range(1, HISTOGRAM_LINES):
            lines.append("  Bin%d  0" % i)
        reply = "\n".join(lines)
    return reply


def _fake_800g_reply_port5_no_module(cmd):
    # Like _fake_800g_reply, but port 5, the third port tested, has no module. As on the
    # real switch, "grep -i SN" then prints only the "Lane SNR" header, 1 line instead of 2.
    reply = _fake_800g_reply(cmd)
    if "grep" in cmd and _port_of(cmd) == 5:
        if cmd.endswith("grep -i SN"):
            reply = "    Lane  SNR(dB)    VEYE(U,M,L)"
        else:
            reply = ""
    return reply


# Which fake switch answers for each template.
FAKE_REPLIES = {
    "200G": _fake_switch_reply,
    "400G": _fake_switch_reply,
    "800G": _fake_800g_reply,
}


def _fake_recording_switch(sent, reply):
    # A fake switch that answers with `reply` and records every command in `sent`.
    def send_command(cmd):
        sent.append(cmd)
        return reply(cmd)

    return _fake_connection(send_command)


def _run_all_steps(key):
    # Runs every step of one template on a fake switch. Returns the workbook and the commands.
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES[key])
    sent = []
    connection = _fake_recording_switch(sent, FAKE_REPLIES[key])
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES[key])
    return workbook, sent


def _count_sent_per_step(sent, expected_starts):
    # Checks each step's commands were sent in order, from its first command and count.
    index = 0
    for i in range(len(expected_starts)):
        first_cmd, count = expected_starts[i]
        assert sent[index] == first_cmd
        index = index + count
    assert index == len(sent)


def _raise_switch_error(cmd):
    raise RuntimeError("switch went away")


def _fake_recording_connection(send_command, calls):
    # Like _fake_connection, but records disconnect() calls in `calls`.
    def disconnect():
        calls.append("disconnect")

    return SimpleNamespace(send_command=send_command,
                           find_prompt=_fake_prompt,
                           disconnect=disconnect)


def _fake_create_connection(connection):
    # Stands in for ssh.create_connection and always returns the same connection.
    def create_connection(device):
        return connection

    return create_connection


def _raise_timeout(cmd):
    raise ReadTimeout("pattern not detected")


def _raise_enable_failure():
    raise ValueError("Failed to enter enable mode")


def _fake_connect_handler_with_failing_enable(calls):
    # Stands in for netmiko.ConnectHandler and records each disconnect() call.
    def disconnect():
        calls.append("disconnect")

    def connect_handler(**device):
        return SimpleNamespace(enable=_raise_enable_failure, disconnect=disconnect)

    return connect_handler


# ---------- excel.format_output ----------

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


def test_format_output_no_output_gets_one_placeholder_line():
    # A port with no module prints nothing. One blank line keeps the later ports in place.
    block = excel.format_output(PROMPT, "show x", [], None, False, 1)
    assert block == [PROMPT + "show x", " "]


def test_format_output_no_output_on_first_port_without_echo_gets_placeholder():
    block = excel.format_output(PROMPT, "show x", [], "skip_first_echo", True, 1)
    assert block == [" "]


def test_format_output_header_with_no_output_raises():
    with pytest.raises(ValueError, match="show int tx"):
        excel.format_output(PROMPT, "show int tx", [], "header", True, 1)


# ---------- excel.paste ----------

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


# ---------- excel.get_cmds_from_excel (real templates) ----------

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


def test_800g_master_tests_the_switch_ports_split_8_ways():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["800G"])
    cmds = excel.get_cmds_from_excel(workbook, "Result-2", "Test command!O3:O34")
    ports = []
    for i in range(len(cmds)):
        ports.append(_port_of(cmds[i]))
    assert ports == PORTS_800G


# ---------- excel.format_output, blank lines for empty ports ----------

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








# ---------- excel.paste_columns ----------

def test_paste_columns_puts_each_block_down_its_own_column_group():
    worksheet = openpyxl.Workbook().active
    excel.paste_columns(worksheet, "B5", [["a1", "a2"], ["b1"], ["c1"]])
    assert worksheet["B5"].value == "a1"
    assert worksheet["B6"].value == "a2"
    assert worksheet["G5"].value == "b1"
    assert worksheet["L5"].value == "c1"
    assert worksheet["C5"].value is None


# ---------- ssh.run ----------

def _reply_invalid_input(cmd):
    return "% Invalid input (at token 2: 'Et99/1')"


def _reply_two_lines(cmd):
    return "line one\nline two"


def test_run_returns_output_as_lines():
    connection = _fake_connection(_reply_two_lines)
    assert ssh.run(connection, "show x") == ["line one", "line two"]


def test_run_switch_error_raises_runtime_error_naming_command():
    # EOS starts the line with "% " when a command fails. That must never be pasted as data.
    connection = _fake_connection(_reply_invalid_input)
    with pytest.raises(RuntimeError, match="Et99/1 phy detail"):
        ssh.run(connection, "show interfaces Et99/1 phy detail")


def test_run_timeout_raises_runtime_error_naming_command():
    connection = _fake_connection(_raise_timeout)
    with pytest.raises(RuntimeError, match="Et5/1"):
        ssh.run(connection, "show interfaces Et5/1 phy detail")


def test_create_connection_enable_failure_disconnects_and_reraises(monkeypatch):
    calls = []
    monkeypatch.setattr(ssh, "ConnectHandler", _fake_connect_handler_with_failing_enable(calls))
    device = {"device_type": "arista_eos", "host": "10.0.0.1"}
    with pytest.raises(ValueError, match="enable"):
        ssh.create_connection(device)
    assert calls == ["disconnect"]


def test_create_connection_skips_enable_on_linux(monkeypatch):
    calls = []
    monkeypatch.setattr(ssh, "ConnectHandler", _fake_connect_handler_with_failing_enable(calls))
    device = {"device_type": "linux", "host": "10.0.0.1"}
    ssh.create_connection(device)   # This would raise if enable() were called.
    assert calls == []


# ---------- main._parse_args ----------

def test_parse_args_returns_known_key(monkeypatch):
    monkeypatch.setattr("sys.argv", ["main.py", "200G"])
    assert main._parse_args() == "200G"


def test_parse_args_unknown_key_prints_usage(monkeypatch):
    monkeypatch.setattr("sys.argv", ["main.py", "300G"])
    with pytest.raises(SystemExit, match="usage"):
        main._parse_args()


def test_parse_args_missing_key_prints_usage(monkeypatch):
    monkeypatch.setattr("sys.argv", ["main.py"])
    with pytest.raises(SystemExit, match="usage"):
        main._parse_args()


# ---------- main._run_on_switch ----------

def test_run_on_switch_closes_connection_after_success(monkeypatch):
    calls = []
    connection = _fake_recording_connection(_fake_switch_reply, calls)
    monkeypatch.setattr(ssh, "create_connection", _fake_create_connection(connection))
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["200G"])
    main._run_on_switch("200G", workbook)
    assert calls == ["disconnect"]


def test_run_on_switch_closes_connection_when_a_step_fails(monkeypatch):
    calls = []
    connection = _fake_recording_connection(_raise_switch_error, calls)
    monkeypatch.setattr(ssh, "create_connection", _fake_create_connection(connection))
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["200G"])
    with pytest.raises(RuntimeError, match="switch went away"):
        main._run_on_switch("200G", workbook)
    assert calls == ["disconnect"]


# ---------- main._save_copy ----------

def test_save_copy_writes_key_and_timestamp_name_in_runs(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    main._save_copy("200G", openpyxl.Workbook())
    saved = []
    for path in (tmp_path / "runs").glob("*.xlsx"):
        saved.append(path)
    assert len(saved) == 1
    assert re.fullmatch(r"200G_\d{8}_\d{6}\.xlsx", saved[0].name)


# ---------- main.run_steps (fake switch, real template) ----------

def test_run_steps_200g_runs_the_six_steps_in_order():
    workbook, sent = _run_all_steps("200G")
    expected_starts = [
        ("show interfaces Et1/1 phy detail | grep -i SN", 32),         # Step-1
        ("clear phy counters et1/1", 32),                               # Step-2
        ("show interfaces Et1/1 phy detail | grep -i Interface", 32),  # Step-3
        ("show interfaces Et1/1 phy detail | grep -i Pre-FEC", 32),    # Step-4
        ("show interfaces transceiver", 1),                             # Step-5
        ("sho int et 1/1", 32),                                         # Step-6
    ]
    _count_sent_per_step(sent, expected_starts)


def test_run_steps_200g_step1_sn_lands_where_column_n_reads():
    # Step-1 formulas read N2 = G3, N3 = G5, ... N33 = G65
    workbook, sent = _run_all_steps("200G")
    worksheet = workbook["Step-1 (SN)"]
    assert worksheet["G2"].value == PROMPT + "show interfaces Et1/1 phy detail | grep -i SN"
    for port in range(1, 33):
        cell = "G" + str(2 * port + 1)
        assert worksheet[cell].value == "  Transceiver SN   SN%02d" % port
    assert worksheet["G66"].value is None


def test_run_steps_200g_empty_port_keeps_later_ports_aligned():
    # Port 5 has no module. Its SN cell, G11, is blank and ports 6 to 32 stay in place.
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["200G"])
    connection = _fake_connection(_fake_reply_port5_empty)
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["200G"])
    for tab in ("Step-1 (SN)", "Step-3 (Link Status)", "Step-4 (Pre-FEC BER)"):
        worksheet = workbook[tab]
        assert worksheet["G11"].value == " "
        assert worksheet["G13"].value.endswith("06")
        assert worksheet["G65"].value.endswith("32")
        assert worksheet["G66"].value is None


def test_run_steps_200g_step2_clears_all_32_ports_and_pastes_nothing():
    workbook, sent = _run_all_steps("200G")
    clears = []
    for i in range(len(sent)):
        if sent[i].startswith("clear"):
            clears.append(sent[i])
    assert len(clears) == 32
    for port in range(1, 33):
        assert clears[port - 1] == "clear phy counters et%d/1" % port
    worksheet = workbook["Step-2 (Clear status)"]
    assert worksheet["C2"].value is None


def test_run_steps_200g_step3_link_status_lands_where_column_o_reads():
    # Step-3 formulas read O2 = G3, O3 = G5, ... O33 = G65
    workbook, sent = _run_all_steps("200G")
    worksheet = workbook["Step-3 (Link Status)"]
    for port in range(1, 33):
        cell = "G" + str(2 * port + 1)
        assert worksheet[cell].value == "  Interface state  up-%02d" % port
    assert worksheet["G66"].value is None


def test_run_steps_200g_step4_pre_fec_lands_where_column_o_reads():
    # Step-4 formulas read O2 = G3, O3 = G5, ... O33 = G65
    workbook, sent = _run_all_steps("200G")
    worksheet = workbook["Step-4 (Pre-FEC BER)"]
    for port in range(1, 33):
        cell = "G" + str(2 * port + 1)
        assert worksheet[cell].value == "  Pre-FEC bit error rate  BER-%02d" % port
    assert worksheet["G66"].value is None


def test_run_steps_200g_step5_dom_starts_at_port_line_under_result_header():
    # As in the 400G sample, D4 says "Result" and the output starts at D5 with "Port".
    workbook, sent = _run_all_steps("200G")
    worksheet = workbook["Step-5 (DOM)"]
    assert worksheet["D4"].value == "Result"
    assert worksheet["D5"].value == TRANSCEIVER_OUTPUT[3]
    assert worksheet["D6"].value == TRANSCEIVER_OUTPUT[4]
    assert worksheet["D7"].value == TRANSCEIVER_OUTPUT[5]
    assert worksheet["D8"].value is None


def test_run_steps_200g_step6_line_protocol_lands_where_column_p_reads():
    # Step-6 formulas read P2 = D2, P3 = D23, ... P33 = D653. Port 1 has no command line.
    # Each later port has one plus 20 lines, so port N starts at row 2 + 21 * (N - 1).
    workbook, sent = _run_all_steps("200G")
    worksheet = workbook["Step-6 (Line protocol)"]
    for port in range(1, 33):
        row = 2 + 21 * (port - 1)
        expected = "Ethernet%d/1 is up, line protocol is up (connected)" % port
        assert worksheet["D" + str(row)].value == expected
    assert worksheet["D23"].value != worksheet["D22"].value
    assert worksheet["D22"].value.startswith(PROMPT + "sho int et 2/1")


# ---------- 400G (DX5) ----------

def test_run_steps_400g_runs_the_seven_steps_in_order():
    workbook, sent = _run_all_steps("400G")
    expected_starts = [
        ("show interfaces Et1/1 phy detail | grep -i SN", 32),                # Step-1
        ("clear phy counters et1/1", 32),                                      # Step-2
        ("sho int sta", 1),                                                    # Step-3
        ("show interfaces Et1/1 phy detail | grep -i Pre-FEC", 32),           # Step-4
        ("show interfaces transceiver", 1),                                    # Step-5
        ("show interfaces Et1/1-$ transceiver dom | grep -i case", 32),       # Step-6
        ("show int et1/1 phy diag error-correction histogram", 32),          # Step-7
    ]
    _count_sent_per_step(sent, expected_starts)


def test_run_steps_400g_step1_sn_lands_where_column_n_reads():
    # Step-1 formulas read N2 = G3, N3 = G5, ... N33 = G65
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-1"]
    assert worksheet["G2"].value == PROMPT + "show interfaces Et1/1 phy detail | grep -i SN"
    for port in range(1, 33):
        cell = "G" + str(2 * port + 1)
        assert worksheet[cell].value == "  Transceiver SN   SN%02d" % port


def test_run_steps_400g_step2_clears_all_32_ports_in_order():
    workbook, sent = _run_all_steps("400G")
    clears = []
    for i in range(len(sent)):
        if sent[i].startswith("clear"):
            clears.append(sent[i])
    assert len(clears) == 32
    for port in range(1, 33):
        assert clears[port - 1] == "clear phy counters et%d/1" % port


def test_run_steps_400g_step3_link_status_one_row_per_port_where_column_o_reads():
    # Step-3 formulas read O2 = G3, O3 = G4, ... O33 = G34, from one "sho int sta" table.
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-3"]
    assert worksheet["G2"].value.startswith("Port ")
    for port in range(1, 33):
        cell = "G" + str(port + 2)
        assert worksheet[cell].value.startswith("Et%d/1 " % port)
        assert worksheet[cell].value.endswith("status-%02d" % port)


def test_run_steps_400g_step4_pre_fec_lands_where_column_o_reads():
    # Step-4 formulas read O2 = G3, O3 = G5, ... O33 = G65
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-4 "]
    for port in range(1, 33):
        cell = "G" + str(2 * port + 1)
        assert worksheet[cell].value == "  Pre-FEC bit error rate  BER-%02d" % port


def test_run_steps_400g_step5_dom_starts_at_port_line_under_result_header():
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-5"]
    assert worksheet["D4"].value == "Result"
    assert worksheet["D5"].value == TRANSCEIVER_OUTPUT[3]
    assert worksheet["D7"].value == TRANSCEIVER_OUTPUT[5]


def test_run_steps_400g_step6_case_temp_lands_where_column_p_reads():
    # Step-6 formulas read P2 = D3, P3 = D5, ... P33 = D65
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-6"]
    assert worksheet["D2"].value.startswith(PROMPT + "show interfaces Et1/1-$ transceiver")
    for port in range(1, 33):
        cell = "D" + str(2 * port + 1)
        assert worksheet[cell].value == "   Case temperature   temp-%02d" % port


def test_run_steps_400g_step7_histogram_first_port_has_no_echo():
    # No formulas read this. As in the sample, port 1 starts at D2 with no command line.
    workbook, sent = _run_all_steps("400G")
    worksheet = workbook["Step-7 "]
    assert worksheet["D2"].value == " Ethernet1/1"
    second_echo = 2 + HISTOGRAM_LINES
    expected = PROMPT + "show int et2/1 phy diag error-correction histogram"
    assert worksheet["D" + str(second_echo)].value == expected
    assert worksheet["D" + str(second_echo + 1)].value == " Ethernet2/1"


def test_run_steps_400g_empty_port_keeps_later_ports_aligned():
    # Port 5 has no module. Its SN, BER and case temperature are blank. Later ports stay put.
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["400G"])
    connection = _fake_connection(_fake_reply_port5_empty)
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["400G"])
    for tab, col in (("Step-1", "G"), ("Step-4 ", "G"), ("Step-6", "D")):
        worksheet = workbook[tab]
        assert worksheet[col + "11"].value == " "
        assert worksheet[col + "13"].value.endswith("06")
        assert worksheet[col + "65"].value.endswith("32")


# ---------- 800G (TH5) ----------

def test_run_steps_800g_runs_the_eight_checks_in_order():
    workbook, sent = _run_all_steps("800G")
    expected_starts = [
        ("clear phy counters et1/1-$", 32),                                  # Clear
        ("show interfaces Et1/1-$ phy detail | grep -i Pre-FEC", 32),        # Result-1 B
        ("show interfaces Et1/1-$ phy detail | grep -i Interface", 32),      # Result-1 F
        ("show interfaces Et1/1-$ transceiver dom | grep -i SNR", 32),       # Result-1 K
        ("show interfaces Et1/1-$ transceiver dom | grep -i case", 32),      # Result-2
        ("show interfaces Et1/1 phy detail | grep -i SN", 32),               # Result-3
        ("show interfaces Et1/1-$ transceiver dom | grep -i optical", 32),   # Result-4
        ("show int et1/1-$ phy diag error-correction histogram", 32),        # Result-5
    ]
    _count_sent_per_step(sent, expected_starts)


def test_run_steps_800g_result1_lanes_line_up_with_port_and_ch_labels():
    # Column A labels each port's block PORT-n, then CH-1 to CH-8. Each port takes 9 rows.
    workbook, sent = _run_all_steps("800G")
    worksheet = workbook["Result-1"]
    for col, text in (("B", "BER"), ("F", "up"), ("K", "dB")):
        for k in range(32):
            port = PORTS_800G[k]
            echo_row = 2 + 9 * k
            assert worksheet[col + str(echo_row)].value.startswith(PROMPT)
            for lane in range(1, 9):
                value = worksheet[col + str(echo_row + lane)].value
                assert value.endswith("%s-%02d-L%d" % (text, port, lane))


def test_run_steps_800g_result2_case_temp_lands_where_summary_f_reads():
    # Summary F2 = Result-2 B3, F3 = B5, ... F33 = B65
    workbook, sent = _run_all_steps("800G")
    worksheet = workbook["Result-2"]
    for k in range(32):
        cell = "B" + str(3 + 2 * k)
        assert worksheet[cell].value == "   Case temperature   temp-%02d" % PORTS_800G[k]


def test_run_steps_800g_result3_sn_lands_where_result5_row2_reads():
    # Summary C reads Result-5 row 2, which reads Result-3 B3, B6 and so on to B96.
    workbook, sent = _run_all_steps("800G")
    worksheet = workbook["Result-3"]
    for k in range(32):
        cell = "B" + str(3 + 3 * k)
        assert worksheet[cell].value == "  Transceiver SN   SN%02d" % PORTS_800G[k]
    assert worksheet["B97"].value == "    Lane  SNR(dB)    VEYE(U,M,L)"
    assert worksheet["B98"].value is None


def test_run_steps_800g_result4_tx_rx_is_16_lines_per_port():
    workbook, sent = _run_all_steps("800G")
    worksheet = workbook["Result-4"]
    for k in range(32):
        echo_row = 2 + 17 * k
        assert worksheet["B" + str(echo_row)].value.startswith(PROMPT)
        assert worksheet["B" + str(echo_row + 16)].value.endswith("-%02d-L16" % PORTS_800G[k])


def test_run_steps_800g_result5_one_histogram_per_column_group_from_row5():
    # Port k's group starts at column B + 5k. Rows 2 to 4 hold the SN formula, the PORT label
    # and the command, and must stay as they are.
    workbook, sent = _run_all_steps("800G")
    worksheet = workbook["Result-5"]
    assert worksheet["B2"].value == "='Result-3'!$B$3"
    assert worksheet["B3"].value == "PORT-1"
    for k in range(32):
        col = get_column_letter(2 + 5 * k)
        port = PORTS_800G[k]
        expected = PROMPT + "show int et%d/1-$ phy diag error-correction histogram" % port
        assert worksheet[col + "5"].value == expected
        assert worksheet[col + "6"].value == " Ethernet%d/1" % port


def test_run_steps_800g_empty_port_keeps_later_ports_aligned():
    # Port 5, the third port tested, has no module and prints only the "Lane SNR" line for
    # its SN. Its SN block becomes 2 blank rows and its lane block 8 blank rows. Every later
    # port stays where the formulas and labels expect it.
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["800G"])
    connection = _fake_connection(_fake_800g_reply_port5_no_module)
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["800G"])
    result3 = workbook["Result-3"]
    assert result3["B9"].value == " "
    assert result3["B10"].value == " "
    assert result3["B12"].value == "  Transceiver SN   SN06"
    result1 = workbook["Result-1"]
    for lane in range(1, 9):
        assert result1["B" + str(20 + lane)].value == " "
    assert result1["B30"].value.endswith("BER-06-L1")


# ---------- replay recorded output from real switches through the real code ----------
#
# The sample workbooks in tests/ hold what a person pasted from a real switch. These tests
# split each pasted column back into each command's output and replay it through
# run_steps into the blank master. Then they follow every Summary formula to the cell it
# reads, and that cell must match the sample. Summary formulas only copy cells, so
# matching cells mean a matching Summary.

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

    return _fake_connection(send_command)


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
    sample = openpyxl.load_workbook("tests/400G_sample.xlsx")
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
    sample = openpyxl.load_workbook("tests/800G_sample.xlsx")
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
