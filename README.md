# RMA

Runs the RMA transceiver test on an Arista switch and fills in a copy of the Excel
template for you. For each tab, it sends the tab's commands over SSH and pastes the output
exactly where you would paste it from PuTTY, so every formula and the Summary work as
before.

## What you need

- Python 3.8 or later
- Network access to the lab switches (the switch details are in `main.py`)
- Modules inserted in the ports you want to test
- Excel, to open the result

## Install

Once, from the `RMA` folder:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If PowerShell blocks `Activate.ps1`, run
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once, then try again.

## Run

From the `RMA` folder, with the virtual environment active (`(.venv)` in the prompt),
choose the data rate:

```powershell
python main.py 200G
python main.py 400G
python main.py 800G
```

You'll see one line per step, then the file it saved:

```
Step-1 (SN) B3:B34: 32 commands
Step-2 (Clear status) B3:B34: 32 commands
...
Step-6 (Line protocol) A2:A33: 32 commands
Saved runs\200G_20261123_164351.xlsx
```

A run takes a few minutes. It clears the switch's PHY counters as one of its steps, as the
manual test does.

## Output

- Each run saves a new file in `runs\`, named `<rate>_<YYYYmmdd_HHMMSS>.xlsx`.
  Rename it as you normally would (for example with the RMA and lot number).
- The templates in `templates\` are only read, never changed.
- **Open the file in Excel and save it once before sharing.** The Summary's formulas are
  calculated when Excel opens the file.
- A port with no module shows as blank in the Summary; every other port stays on its
  own row.

## If something goes wrong

If a run stops, nothing is saved. Fix the cause and run it again from the start.

| Message | What it means | What to do |
| --- | --- | --- |
| `TCP connection to device failed` | The switch didn't answer. | Check you're on the lab network and the switch is up (`ping` its address). |
| `Authentication to device failed` | The username or password was refused. | Check the switch's entry in `SWITCHES` in `main.py`. |
| `Timed out on: <command>` | The switch took longer than 10 s to answer that command. | Check the switch is still reachable, then run again. |
| `Switch error on: <command> -> % ...` | The switch rejected the command. | Check the command in the template tab named in the message. |
| `Header not found in output of: <command>` | A table command (DOM, `sho int sta`) returned no table. | Check the switch and the modules, then run again. |
| `usage: python main.py 200G\|400G\|800G` | The data rate was missing or misspelled. | Run it again with `200G`, `400G` or `800G`. |
| `FileNotFoundError: ... templates/...` | It was started outside the `RMA` folder. | `cd` into the `RMA` folder and run it again. |

## Tests

```powershell
pytest
```

The tests use a fake switch and recorded output from real switches, so they need no lab
access. They run in a few seconds; run them after any change.

## Project layout

| File | Purpose |
| --- | --- |
| `main.py` | Entry point; switch details (`SWITCHES`) and the run order |
| `templates.py` | Each template's file and, per tab, where commands are read and output goes |
| `excel.py` | Reads commands from the workbook, shapes the output, pastes it |
| `ssh.py` | Connects to the switch and runs one command |
| `test_rma.py` | Tests; `tests\` holds the recorded sample workbooks |
| `CLAUDE.md` | Design rules, paste-format rules and template details, for changing the code |
