# Tests for main.py, except run_steps, which has its own files. No switch is needed.

import os
import re
from types import SimpleNamespace

import openpyxl
import pytest

import main
import ssh
import templates
from fakes import fake_prompt, fake_switch_reply


def _raise_switch_error(cmd):
    raise RuntimeError("switch went away")


def _fake_recording_connection(send_command, calls):
    # Like fakes.fake_connection, but records disconnect() calls in `calls`.
    def disconnect():
        calls.append("disconnect")

    return SimpleNamespace(send_command=send_command,
                           find_prompt=fake_prompt,
                           disconnect=disconnect)


def _fake_create_connection(connection):
    # Stands in for ssh.create_connection and always returns the same connection.
    def create_connection(device):
        return connection

    return create_connection


def _skip_switch(template_key, workbook):
    pass   # Stands in for main._run_on_switch, so no switch is needed.


# ---------- _parse_args ----------

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


# ---------- _run_on_switch ----------

def test_run_on_switch_closes_connection_after_success(monkeypatch):
    calls = []
    connection = _fake_recording_connection(fake_switch_reply, calls)
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


# ---------- _save_copy ----------

def test_save_copy_writes_key_and_timestamp_name_in_results_folder(monkeypatch, tmp_path):
    results = tmp_path / "RMA results"
    monkeypatch.setattr(main, "RESULTS_FOLDER", results)
    out = main._save_copy("200G", openpyxl.Workbook())
    saved = []
    for path in results.glob("*.xlsx"):
        saved.append(path)
    assert saved == [out]
    assert re.fullmatch(r"200G_\d{8}_\d{6}\.xlsx", out.name)


# ---------- main ----------

def test_main_opens_the_saved_copy(monkeypatch, tmp_path):
    opened = []
    monkeypatch.setattr("sys.argv", ["main.py", "200G"])
    monkeypatch.setattr(main, "_run_on_switch", _skip_switch)
    monkeypatch.setattr(main, "RESULTS_FOLDER", tmp_path)
    monkeypatch.setattr(os, "startfile", opened.append)   # Don't open Excel during tests.
    main.main()
    assert len(opened) == 1
    assert opened[0].parent == tmp_path
