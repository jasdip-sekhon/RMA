# Tests for ssh.py. Netmiko is faked, so no switch is needed.

from types import SimpleNamespace

import pytest
from netmiko import ReadTimeout   # Tests may import netmiko, to fake its errors.

import ssh
from fakes import fake_connection


def _reply_invalid_input(cmd):
    return "% Invalid input (at token 2: 'Et99/1')"


def _reply_two_lines(cmd):
    return "line one\nline two"


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


# ---------- run ----------

def test_run_returns_output_as_lines():
    connection = fake_connection(_reply_two_lines)
    assert ssh.run(connection, "show x") == ["line one", "line two"]


def test_run_switch_error_raises_runtime_error_naming_command():
    # EOS starts the line with "% " when a command fails. That must never be pasted as data.
    connection = fake_connection(_reply_invalid_input)
    with pytest.raises(RuntimeError, match="Et99/1 phy detail"):
        ssh.run(connection, "show interfaces Et99/1 phy detail")


def test_run_timeout_raises_runtime_error_naming_command():
    connection = fake_connection(_raise_timeout)
    with pytest.raises(RuntimeError, match="Et5/1"):
        ssh.run(connection, "show interfaces Et5/1 phy detail")


# ---------- create_connection ----------

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
