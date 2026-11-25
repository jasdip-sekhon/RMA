# CLAUDE.md — RMA

## What this is

Automates the manual RMA test loop for optical transceivers on Arista EOS switches.
Today an engineer goes through an Excel template tab by tab: copy the tab's commands,
paste them into an SSH session (PuTTY), copy the output, and paste it back into the tab.
Formulas in the workbook (Summary, SN/Link/BER/Temp columns) then pick values out of the
pasted output **by fixed row position**.

This tool does the same thing: open SSH, run each tab's commands in order, and paste the
output into a copy of the template **exactly where a human would paste it**, so every
existing formula keeps working. It is a straightforward, sequential, deterministic script.
Keep it that way: no frameworks, no plugin layers, no parsing/verdict logic unless asked.

One switch per data rate:

| Key  | Platform         | Template file (starts with)                       |
| ---- | ---------------- | ------------------------------------------------- |
| 200G | TH4 (7060DX4-32) | `TH4_Arista_command_for_200G_2-FR4_...TEMPLATE` |
| 400G | DX5 (7060DX5-32) | `DX5_Arista_command_for_400G_FR4_...TEMPLATE`   |
| 800G | TH5              | `TH5_Arista_command_for_800G_...TEMPLATE`       |

There is also a SONiC box (`device_type: linux`) used for SSH testing.

## Requirements

### Functional — what the tool must do

- **One SSH session per run.** Connect once, run every step, disconnect in `finally`.
  No per-command connect/disconnect.
- **Output identical to a manual paste.** Same rows, same echo lines, same order as a
  PuTTY paste, so the workbook's fixed-position formulas keep working (see
  *Paste format rules*).
- **Never overwrite the template.** Save a new `<key>_<YYYYmmdd_HHMMSS>.xlsx` (e.g.
  `200G_20261123_143005.xlsx`) in `RESULTS_FOLDER` (`Documents\RMA results`), then open it
  in Excel; the masters in `templates/` are only ever read.
- **Fail loudly, no retries.** Any connection, enable, or command failure raises an
  ordinary exception (naming the host or command) that stops the run with a traceback.
  Helpers never call `SystemExit`; only `main.py` decides how the run ends. Never retry,
  never skip and carry on.
- **Support all three templates:** 200G (TH4), 400G (DX5), 800G (TH5), chosen on the
  command line (`python main.py 200G`); the file comes from `TEMPLATE_FILES`.

### Environment

- Python 3.10+ (pytest 9 and current netmiko 4.x need it)
- netmiko 4.x (needed for `ReadTimeout`)
- openpyxl 3.1+
- SSH access to the lab switches
- Excel, to open and save the output once (recalculates formulas) before sharing

## Architecture

### Run flow

```
python main.py 200G | 400G | 800G
        │
        ▼
_parse_args()                        key; anything else → usage message (SystemExit)
        │
        ▼
load_workbook(TEMPLATE_FILES[key])   formulas kept; original never written
        │
        ▼
create_connection(SWITCHES[key])     one session for the whole run
get_prompt(connection)               e.g. "7060DX4-32-F#"
        │
        ▼
run_steps: for each step (a Step dataclass) in TEMPLATES[key]:       in run order
    cmds = get_cmds_from_excel(wb, step.tab, step.commands)    commands may name a tab
    for j in range(len(cmds)):
        lines = run(connection, cmds[j])      timeout → RuntimeError naming the command
        blocks.append(format_output(prompt, cmds[j], lines, step.special, j == 0,
                                    step.lines_per_port))
    paste None        → nothing pasted (run only)
    special "columns" → paste_columns(wb[step.tab], step.paste, blocks)
    otherwise         → paste(wb[step.tab], step.paste, all blocks joined)
        │
        ▼
ssh.close(connection)                in finally
        │
        ▼
save RESULTS_FOLDER/<key>_<YYYYmmdd_HHMMSS>.xlsx
        │
        ▼
os.startfile(out)                    opens it in Excel, which must save it once
```

### File boundaries

| File             | Knows about                                                                | Must never touch                                               |
| ---------------- | -------------------------------------------------------------------------- | -------------------------------------------------------------- |
| `main.py`      | CLI args, `SWITCHES`, run order, opening and saving the workbook         | netmiko directly, cell positions or paste formatting           |
| `templates.py` | each template's file (`TEMPLATE_FILES`) and step table (`TEMPLATES`)       | openpyxl, SSH, the disk; it is plain data only                 |
| `excel.py`     | openpyxl workbook/worksheet objects: reading commands, formatting, pasting | SSH, the disk (never opens or saves files), choosing a template |
| `ssh.py`       | netmiko: connecting, enable, getting the prompt, running one command, closing | Excel, templates, the disk, retries                         |

Only `ssh.py` imports netmiko. Only `main.py` opens or saves files.

### Import direction

```
main.py ──► templates.py
        ├─► excel.py
        └─► ssh.py
```

Imports run one way. `templates.py`, `excel.py` and `ssh.py` never import each other or
anything from `main.py`. Because each of the three depends on nothing else in the
project, each can be tested on its own: `excel.py` with an in-memory workbook,
`templates.py` by reading its rows against the workbook, `ssh.py` against one switch, and `main.py`
with a fake connection that replays recorded output (see *Testing*). A change in one
file never has to be chased backwards into another.

## Coding rules (from the user — follow them)

- **One purpose per file.** *(Single Responsibility Principle; high cohesion)*
- **Imports run one way**, so a change is never chased backwards:
  `main -> {templates, excel, ssh}`. None of those three import each other or
  anything back from `main`. Only `ssh.py` imports netmiko. *(Acyclic Dependencies
  Principle; layered architecture)*
- **Before splitting** a group of code into its own file, check the group calls nothing
  outside itself. Split only when that holds (then the split is mechanical); otherwise don't.
  *(low coupling)*
- The user prefers simple and explicit over clever. *(KISS; "Explicit is better than
  implicit", Zen of Python)* Propose, don't add, extra features. *(YAGNI)*
- The user is reviewing the code **one file at a time**. Work on the file being discussed;
  don't rewrite other files unless asked.

### Style

- **C-like Python.** Index loops (`for i in range(len(x))`), plain `if` blocks, `while`
  loops, sentinel values, string concatenation. No list comprehensions, for/else,
  recursion, `next(...)`, walrus, or several nested calls on one line. One short
  `a or (b and c)` condition is fine when it reads as one sentence. *(house style; no
  common name)*
- **Short functions, one job each.** Short and plain beats long and spelled out; if a
  function feels long, look for a simpler shape before adding helpers. *(Single
  Responsibility Principle, at function level)*
- **Plain functions and data.** No frameworks or abstraction layers, and no classes with
  behaviour. The one kind of class allowed is a data-only `@dataclass(frozen=True)` with
  no methods, like a C++ struct. Fixed facts (steps, switches, template files) live in
  tables, not in `if template == ...` branches. A step row is a `Step` dataclass, read as
  `step.tab`, `step.paste`, etc. Write optional types as `Optional[str]`, as the code
  already does. *(table-driven methods / data-driven design; Unix Rule of Representation)*
- **Each file imports what it uses itself.** Python imports are not inherited; never
  reach through another module (`excel.openpyxl`). Drop imports nothing uses.
  *(Include What You Use)*
- **Follow PEP 8 naming and layout.** `snake_case` functions and variables, `UPPER_CASE`
  module-level constants (`SWITCHES`, `TEMPLATES`), leading `_` for helpers used only in
  their own file, 4-space indents, two blank lines between top-level functions.
  *(PEP 8)*
- **Comments explain why, not what.** See *Writing* below.

### Safety rules (MISRA C principles that carry over to Python)

- **No recursion.** *(MISRA 17.2)*
- **Every loop is bounded:** it runs over a fixed list or range, or a `while` whose index
  only grows toward a fixed limit. *(MISRA Dir 4.1 intent)*
- **Every variable is set before it is used,** on every path. *(MISRA 9.1)*
- **Explicit comparisons:** `if x is not None`, `if special == "header"`, `while i < n`;
  don't rely on truthiness of strings, lists, or numbers. *(MISRA 10.x / 14.4 intent)*
- **One statement per line; no side effects inside conditions.** *(MISRA 13.x)*
- **Single point of exit:** one `return`, at the end of the function. Set a result
  variable in each branch and return it once. Raising an exception is allowed anywhere.
  *(MISRA 15.5)*
- **No unused code or data:** no dead branches, unused functions, variables, imports, or
  table entries. *(MISRA 2.2, 2.3)*
- **Every error is handled or raised:** never ignore a failure or a return value that
  signals one. *(MISRA Dir 4.7, 17.7)*

### Design

- **`main()` reads as an outline**: parse args, load, run on the switch, save. Details go
  in named helpers; don't add a helper that only forwards a call unless a boundary rule
  requires it. *(stepwise refinement / top-down design; Stepdown Rule; Composed Method)*
- **Netmiko lives only in `ssh.py`.** `main.py` never calls a Netmiko method directly,
  even on the `connection` object it holds; add a small `ssh` function instead (e.g.
  `ssh.get_prompt`). `main.py` treats `connection` as an opaque handle. *(information
  hiding; wrapping a third-party library / Adapter pattern)*
- **Errors: helpers raise ordinary exceptions** naming the host or command
  (`RuntimeError`, `ValueError`, or netmiko's own). Only `main.py` may use `SystemExit`,
  e.g. for the usage message. No retries. *(fail fast)*
- **No defensive checks for things that can't happen.** The templates are fixed, so don't
  validate tab names, blank cells, or formula shapes. Check only what can change at run
  time, such as switch output (e.g. the missing `Port` header). *(validate at the
  boundaries; YAGNI)*
- **Testable without a switch.** Functions that talk to the switch take the connection as
  an argument (`run_steps(connection, ...)`) so tests can pass a fake. *(dependency
  injection; test doubles / fakes)*
- **Design patterns in use** (keep to these; don't add new ones without asking):
  table-driven design (`TEMPLATES`, `TEMPLATE_FILES`, `SWITCHES`), adapter/wrapper
  (`ssh.py` around Netmiko), dependency injection (`run_steps(connection, ...)`), a
  pipeline (read commands → run → format → paste), and `main()` as an outline.
- **Each fact lives in one place.** The data-rate keys (`200G`, `400G`, `800G`) appear in
  three tables: `SWITCHES` in `main.py`, and `TEMPLATE_FILES` and `TEMPLATES` in
  `templates.py`. Adding or renaming a rate means updating all three. Don't copy a fact
  (a cell, a range, a host) into a second place. *(single source of truth; DRY)*
- **Pin dependency versions** in `requirements.txt` to the major versions known to work
  (`netmiko>=4,<5`, `openpyxl>=3.1,<4`), so a fresh `pip install` can't pull a breaking
  release. *(pinned dependencies)*

### Writing code comments

- **Comment only what the code can't say.** That means why something is done, which
  formula reads a cell, or a surprise in the switch output. If a comment repeats the
  code, delete it. *(Ousterhout, A Philosophy of Software Design, ch. 13)*
- **Block comments are complete sentences. Inline comments can be short labels**, such
  as `# Pre-FEC BER` or `# N2 = G3 .. G65` in a table. Keep comments up to date with the
  code. *(PEP 8, Comments)*
- **A small diagram is fine** when it shows something more clearly than words, like the
  prompt and command layout above `format_output`.
- **Few comments.** Don't spend effort polishing the wording of a comment. Spend it on
  whether the comment is needed at all.

### Writing the README and docs

- **Plain sentences.** Short, everyday sentences in the second person and present tense.
  Avoid semicolons, colon chains and clipped labels like "Note:". Read it over once
  against the 7 Cs (clear, concise, concrete, correct, coherent, complete, courteous) as
  a final check. *(Google developer documentation style guide; plain language, ISO 24495-1)*
- **The README is for the people who run the tool.** Order it the way they work through
  it: what it is, what you need, setting up a new PC, running it, the output, what to do
  when something goes wrong, tests, then the file layout. Design detail stays in this
  file. *(README best practice)*
- **Concrete and correct.** Use real commands, file names and error messages, and check
  each one against the code or a real run before writing it down.
- **No secrets, and no facts copied from the code.** No passwords. Point to where a fact
  lives (`SWITCHES` in `main.py`) instead of repeating it, so it can't go stale.
- **Keep the README current.** Any change to how the tool is set up or run, what it
  saves, or what an error message says needs the README updated in the same commit.

## Layout

```
RMA/
├── main.py          # SWITCHES, RESULTS_FOLDER + main(): open template, connect, run steps, save copy, open it
├── templates.py     # per template: its .xlsx file, where each tab's commands are, where output goes
├── excel.py         # read commands from the workbook, shape output, paste it back
├── ssh.py           # netmiko connection + run one command
├── requirements.txt # netmiko, openpyxl, pytest
├── pytest.ini       # tells pytest the tests are in tests/ and the code is in this folder
├── templates/       # blank .xlsx masters (never written to by a run)
└── tests/
    ├── fakes.py         # fake switch pieces shared by more than one test file
    ├── test_excel.py    # one test file per module
    ├── test_ssh.py
    ├── test_templates.py
    ├── test_main.py     # main.py, except run_steps
    ├── test_run_steps.py  # every template's steps on a fake switch
    ├── test_replay.py   # recorded real output replayed through run_steps
    └── data/            # 400G_sample.xlsx, 800G_sample.xlsx (masters with sample output)
```

Results go outside the repo, to `RESULTS_FOLDER` in `main.py` (`Documents\RMA results`).

## Paste format rules (why the formulas work)

The output must look like a PuTTY paste, row for row; formulas read it by fixed row.

- Default: each command contributes an echo line (`prompt + command`, rebuilt because
  netmiko strips it) followed by its output.
- `"skip_first_echo"` (200G Step-6, 400G Step-7): no echo line for the first port.
  200G Step-6 formulas read every 21st row (P2 = D2, P3 = D23, …), so `sho int` must print
  exactly 20 lines per port (confirmed on the real 200G switch).
- `"header"` (Step-5, 400G Step-3): no echo line; output starts at the `Port …` line.
- `"columns"` (800G Result-5): one port per column group, 5 columns apart, each starting at
  row 5 (B5, G5, L5 … FA5) with echo + histogram. Rows 2–4 of each group (SN formula,
  `PORT-n` label, command text) are never written.
- 800G reads every command from `Test command` (`commands` like `"Test command!D3:D34"`);
  each `EtN/1-$` command covers 8 lanes. Lines per port: Result-1 8 (matches its `CH-1..8`
  labels), Result-2 1 (Summary F reads B3, B5 … B65), Result-3 **2** (`grep -i SN` also
  matches the `Lane SNR(dB)` line; Result-5 row 2 reads B3, B6 … B96 for the Summary's SN),
  Result-4 16 (TX + RX per lane).
- Never write a cell whose value starts with `=`.
- **Never add or drop a line inside a block:** one missing/extra line shifts every later
  port, and the Summary silently pairs results with the wrong SN.
- **`lines_per_port` is exact.** It defaults to 1. A step whose output can be any length
  (a histogram) opts out with `None`; `header` steps are never checked. Only 400G Step-7
  and 800G Result-5 use `None`, and 200G Step-6 states its 20.
- **Too many lines stop the run** with a `ValueError` naming the command and both counts.
  One extra line would shift every later port, and the Summary would pair results with
  the wrong SN. New rows are checked unless someone deliberately writes `None`.
- **One exception, empty ports:** a command that prints fewer lines than the step's
  `lines_per_port` (1 for `None` steps) is treated as an empty port: its whole block
  becomes that many placeholder lines `" "`, so later ports stay aligned and the Summary
  shows a blank for that port. This covers "nothing printed" and partial output (no
  module on 800G: `grep -i SN` prints only the `Lane SNR` line, 1 of 2); the partial
  lines are dropped. It is `" "`, not `""`: openpyxl saves `""` as an empty cell, which
  Excel formulas show as `0`. Doesn't apply to `"header"` steps (missing header → error).
- **Switch errors stop the run:** `ssh.run` raises `RuntimeError` if any output line
  starts with `% ` (how EOS reports a failed command, e.g. `% Invalid input`), so an
  error is never pasted as data or blanked as an empty port.

## Known template issues (flag, don't silently fix)

- **Fixed:** 200G and 400G clear lists skipped ports 5 and 25 (`et4/1`, `et24/1`
  duplicated) in `Step-2` and `META RMA CMD` col D. Corrected in the masters by editing
  only those cells' XML.
- **Fixed:** the 400G master held sample output in its paste columns (Step-1/3/4 col G,
  Step-5/6/7 col D, up to row 673), which `paste` (no clearing) would have left below
  shorter real output. Blanked by editing only those cells' XML; formulas, formatting and
  every other part unchanged. The original, with the sample, is `tests/data/400G_sample.xlsx`.
  All masters are now blank in their paste columns; keep them that way.
- **Not a typo:** the 800G port list goes `…17, 19, 21…` on purpose. On the lab switch,
  port 18 is a single 800G port (`Et18/2`–`8` inactive; the user thinks it may be
  damaged), while every tested port, 19 included, is split 8 × 100G. Changing 18 in
  briefly broke Result-1 (1 lane instead of 8) and was reverted. If port 18 is repaired
  and re-split, `Test command` row 12 and Result-5 `AU3`/`AU4` are the 10 cells to change.
- **Fixed:** the 800G master's Result tabs held sample output; blanked like 400G's
  (Result-5 rows 2–4 kept). The original is `tests/data/800G_sample.xlsx`.
- 800G `Test command` col H alternates `grep -i Interface` / `grep -i  Interface` (one or
  two spaces) by row; harmless to EOS, sent verbatim.
- `Et1/1-$` in TH5/DX5 commands is a native EOS range (one command returns all lanes);
  send verbatim.
- TH4 Summary has ~15,600 junk `=#REF!` cells far right (col YD+); harmless.
- openpyxl drops cached formula values and calcChain on save; workbook must be opened
  and saved once in Excel before sharing. Images/formatting survive (verified).

## Testing

**Unit tests:** in `tests/`, one file per module, run from the `RMA` folder with `pytest`
(install once with `pip install -r requirements.txt`). Plain test functions and `assert`,
no classes; one behaviour per test, named after it; fakes are plain functions in a
`SimpleNamespace`. A helper used by one test file stays in that file. Only helpers used by
several files go in `tests/fakes.py`, which test files import by name (`from fakes import
...`). No `conftest.py`, so every helper a test uses is visible in its imports.
Tests may import netmiko (to fake its errors); app code still may not, outside `ssh.py`.
Add a `run_steps` test for each new row in `TEMPLATES`. Run `pytest` after every change.

**Test first (TDD: red → green → refactor).** For any new behaviour or template row:
1. **Red:** write the test first (e.g. the cells a new step's formulas read, found by
   inspecting the tab) and run it; it must fail for the expected reason.
2. **Green:** write the smallest change that makes it pass (e.g. add the `TEMPLATES` row).
3. **Refactor:** tidy the code with all tests still passing.

Every function with behaviour of its own has a test, including error paths (e.g. enable
failure disconnects and re-raises; a bad argument prints the usage message). Prove a new
test can fail before trusting it.

**Replay tests** (in `tests/test_replay.py`, part of every `pytest` run): real output a person
pasted from a real switch is split back into (command → output lines) at the echo lines,
replayed through `run_steps` into the blank master by a fake switch, and every Summary
cell is followed through its formulas to the cell it finally reads, which must equal the
sample. Summary formulas only copy cells, so equal inputs mean an equal Summary; no Excel
recalculation is needed.

- 400G: `tests/data/400G_sample.xlsx`, Summary C, F, G, H, I → **0 mismatches of 160**.
- 800G: `tests/data/800G_sample.xlsx` (Result-2, Result-3), Summary C, F, G → **0 of 96**.
  Confirms on real output: 400G case temp is 1 line per port; 800G SN is 2 lines
  (`Transceiver SN` + `Lane SNR`) per port with a module.
- 200G: no recorded workbook in the project (`200G_FR4_158pcs_RMAIW20260907001_LOT1_
  initial.xlsx` is elsewhere), so no replay test; the real 200G run filled all 32 ports.
- Still unverified with modules present: 800G Result-1 8 lanes and Result-4 16 lines on
  every port (the sample shows them, but those tabs are read by people only).
