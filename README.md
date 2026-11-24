# RMA

Runs the RMA transceiver test on an Arista switch and fills in a copy of the Excel
template for you. For each tab, it sends the tab's commands over SSH and pastes the output
exactly where you would paste it from PuTTY, so every formula and the Summary work as
before.

## What you need

- Python 3.8 or later
- Git for Windows, to get the code

## Set up on a new PC

Do this once on each PC.

1. **Install Python** from [python.org](https://www.python.org/downloads/). In the
   installer, tick **Add python.exe to PATH**.
2. **Install Git for Windows** from [git-scm.com](https://git-scm.com/download/win).
3. **Get the code.** Open PowerShell and run:

   ```powershell
   cd $HOME\Documents
   git clone https://github.com/jasdip-sekhon/RMA.git
   cd RMA
   ```

   If the repository is private, Git opens a browser window the first time and asks you
   to sign in to GitHub.
4. **Install the Python packages** into the project's own environment:

   ```powershell
   python -m venv .venv
   .venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

   This creates the `.venv` folder that the `Run ... .bat` files use. It isn't in Git,
   because it only works on the PC it was made on. Don't copy it from another PC.
5. **Check it works.** This needs no lab access:

   ```powershell
   .venv\Scripts\python.exe -m pytest
   ```

   Every test should pass.
6. **Connect to the lab network** (`Innolight-Lab`) and check the switch answers, for
   example `ping 192.168.11.220`.

### Getting updates

From the `RMA` folder, run `git pull`. If `requirements.txt` changed, run the
`pip install` line from step 4 again.

## Run

In the `RMA` folder, double-click the file for the data rate you're testing:

- `Run 200G.bat`
- `Run 400G.bat`
- `Run 800G.bat`

A window opens and shows one line per step, then the file it saved:

```
Step-1 (SN) B3:B34: 32 commands
Step-2 (Clear status) B3:B34: 32 commands
...
Step-6 (Line protocol) A2:A33: 32 commands
Saved runs\200G_20261123_164351.xlsx

Done. Open the file in Excel and save it once before sharing.
```

The window stays open until you press a key, so you can read the result. A run takes a
few minutes. It clears the switch's PHY counters as one of its steps, as the manual test
does.

You can also run it from a terminal in the `RMA` folder:

```powershell
.venv\Scripts\python.exe main.py 200G
```

## Output

- Each run saves a new file in `runs\`, named `<rate>_<YYYYmmdd_HHMMSS>.xlsx`.
  Rename it as you normally would (for example with the RMA and lot number).
- The templates in `templates\` are only read, never changed.
- **Open the file in Excel and save it once before sharing.** The Summary's formulas are
  calculated when Excel opens the file.
- A port with no module shows as blank in the Summary. Every other port stays on its
  own row.

## If something goes wrong

If a run stops, nothing is saved. Fix the cause and run it again from the start.

| Message                                                | What it means                                                                                     | What to do                                                                                                |
| ------------------------------------------------------ | ------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `TCP connection to device failed`                    | The switch didn't answer.                                                                         | Check you're on the lab network and the switch is up (`ping` its address).                              |
| `Authentication to device failed`                    | The username or password was refused.                                                             | Check the switch's entry in `SWITCHES` in `main.py`.                                                  |
| `Timed out on: <command>`                            | The switch took longer than 10 s to answer that command.                                          | Check the switch is still reachable, then run again.                                                      |
| `Switch error on: <command> -> % ...`                | The switch rejected the command.                                                                  | Check the command in the template tab named in the message.                                               |
| `Header not found in output of: <command>`           | A table command (DOM, `sho int sta`) returned no table.                                         | Check the switch and the modules, then run again.                                                         |
| `Expected 1 line(s) per port, got 2 from: <command>` | A port printed more lines than the template expects, so its results would land on the wrong rows. | Don't run it again. Report it, because the switch output has changed and `templates.py` needs updating. |
| `usage: python main.py 200G\|400G\|800G`               | The data rate was missing or misspelled.                                                          | Run it again with `200G`, `400G` or `800G`.                                                         |
| `FileNotFoundError: ... templates/...`               | It was started from a terminal outside the `RMA` folder.                                        | Use the `Run ... .bat` files, or `cd` into the `RMA` folder first.                                  |

## Tests

From the `RMA` folder:

```powershell
.venv\Scripts\python.exe -m pytest
```

The tests use a fake switch and recorded output from real switches, so they need no lab
access. They take a few seconds. Run them after any change.

## Project layout

| File                                                 | Purpose                                                                      |
| ---------------------------------------------------- | ---------------------------------------------------------------------------- |
| `Run 200G.bat`, `Run 400G.bat`, `Run 800G.bat` | Double-click launchers                                                       |
| `main.py`                                          | Entry point, with the switch details (`SWITCHES`) and the run order        |
| `templates.py`                                     | Each template's file and, per tab, where commands are read and output goes   |
| `excel.py`                                         | Reads commands from the workbook, shapes the output, pastes it               |
| `ssh.py`                                           | Connects to the switch and runs one command                                  |
| `test_rma.py`                                      | Tests. The `tests\` folder holds the recorded sample workbooks             |
| `CLAUDE.md`                                        | Design rules, paste-format rules and template details, for changing the code |
