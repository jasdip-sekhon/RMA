# Tests for main.run_steps. A fake switch answers each real template's commands, and each
# test checks the cells that template's formulas read.

import re

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

import main
import templates
from fakes import (HISTOGRAM_LINES, PORTS_800G, PROMPT, TRANSCEIVER_OUTPUT,
                   fake_connection, fake_switch_reply, port_of)


def _fake_reply_port5_empty(cmd):
    # Like fake_switch_reply, but port 5 has no module, so its grep commands print nothing.
    reply = fake_switch_reply(cmd)
    if "grep" in cmd and port_of(cmd) == 5:
        reply = ""
    return reply


def _fake_reply_port7_extra_sn_line(cmd):
    # Like fake_switch_reply, but port 7's SN grep matches one line too many.
    reply = fake_switch_reply(cmd)
    if cmd.endswith("grep -i SN") and port_of(cmd) == 7:
        reply = reply + "\n  Another line that matches SN"
    return reply


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
        port = port_of(cmd)
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
    if "grep" in cmd and port_of(cmd) == 5:
        if cmd.endswith("grep -i SN"):
            reply = "    Lane  SNR(dB)    VEYE(U,M,L)"
        else:
            reply = ""
    return reply


# Which fake switch answers for each template.
FAKE_REPLIES = {
    "200G": fake_switch_reply,
    "400G": fake_switch_reply,
    "800G": _fake_800g_reply,
}


def _fake_recording_switch(sent, reply):
    # A fake switch that answers with `reply` and records every command in `sent`.
    def send_command(cmd):
        sent.append(cmd)
        return reply(cmd)

    return fake_connection(send_command)


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


# ---------- 200G (TH4) ----------

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
    connection = fake_connection(_fake_reply_port5_empty)
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["200G"])
    for tab in ("Step-1 (SN)", "Step-3 (Link Status)", "Step-4 (Pre-FEC BER)"):
        worksheet = workbook[tab]
        assert worksheet["G11"].value == " "
        assert worksheet["G13"].value.endswith("06")
        assert worksheet["G65"].value.endswith("32")
        assert worksheet["G66"].value is None


def test_run_steps_200g_extra_output_line_stops_the_run():
    workbook = openpyxl.load_workbook(templates.TEMPLATE_FILES["200G"])
    connection = fake_connection(_fake_reply_port7_extra_sn_line)
    with pytest.raises(ValueError, match="Et7/1"):
        main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["200G"])


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
    connection = fake_connection(_fake_reply_port5_empty)
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
    connection = fake_connection(_fake_800g_reply_port5_no_module)
    main.run_steps(connection, workbook, PROMPT, templates.TEMPLATES["800G"])
    result3 = workbook["Result-3"]
    assert result3["B9"].value == " "
    assert result3["B10"].value == " "
    assert result3["B12"].value == "  Transceiver SN   SN06"
    result1 = workbook["Result-1"]
    for lane in range(1, 9):
        assert result1["B" + str(20 + lane)].value == " "
    assert result1["B30"].value.endswith("BER-06-L1")
