# DataTester

A Claude Code skill for generic Excel data validation and classification. It guides you through a structured 5-phase workflow to analyze input fields, collect judgment rules, run a demo, and produce a fully-processed output Excel file — without hardcoding any business logic.

## What it does

Given an input Excel file and your judgment rules, DataTester will:

1. **Analyze** your input file's column structure and align with you on what each field means
2. **Collect rules** — either described in natural language, or loaded from a structured rule template (`DataTester判断规则模板.xlsx`)
3. **Analyze** your output template's column structure and formatting requirements
4. **Demo run** — process the first 20 rows so you can verify the logic before committing
5. **Full run** — process all rows in batches of 20 with triple verification per batch, then merge and export

Sub-schemes (named configurations) can be saved after a successful demo and reloaded in future sessions.

## Requirements

- [Claude Code](https://claude.ai/code) (CLI, desktop app, or IDE extension)
- Python 3 with `openpyxl` installed:
  ```bash
  pip install openpyxl
  ```

## Installation

**macOS / Linux**

```bash
cp -r DataTester ~/.claude/skills/DataTester
```

**Windows** (PowerShell)

```powershell
Copy-Item -Recurse DataTester "$env:USERPROFILE\.claude\skills\DataTester"
```

The skill will be automatically loaded by Claude Code on the next session.

> **Windows users:** Python is usually invoked as `python` or `py` rather than `python3`. If a command fails, try replacing `python3` with `python`. The Python scripts themselves are fully cross-platform.

## Usage

Trigger the skill by mentioning one of these keywords in your conversation with Claude:

> 数据测试、DataTester、数据校验、Excel校验、字段校验、数据判断、数据处理、分析Excel

Claude will then walk you through the 5-phase workflow interactively.

### Starting options

When the skill launches, it asks:

1. **Create new scheme** — walk through all 5 phases to configure from scratch
2. **Use existing sub-scheme** — load a previously saved configuration (skips phases 1–3)

### Judgment rules: two input methods

**Option A — Natural language description**
Describe your rules in plain text directly in the chat. Claude will parse and confirm its understanding before proceeding.

**Option B — Rule template file**
Use the structured Excel template (`DataTester判断规则模板.xlsx`). To generate a blank template to your Desktop:

```bash
python3 ~/.claude/skills/DataTester/gen_template.py
```

Fill in the template with your rules, then provide the file path when prompted.

The template has 5 columns:

| Column | Description |
|--------|-------------|
| 适用字段 | Input column identifier (e.g. A列, C列) |
| 判断条件 | Condition in natural language |
| 输出字段 | Output column name to write to |
| 输出值 | Value to write when condition is met |
| 优先级 | Priority (lower number = higher priority) |
| 备注 | Notes |

## File structure

```
DataTester/
├── SKILL.md              # Skill definition — read by Claude Code
├── read_excel.py         # Reads an Excel file range, outputs JSON
├── write_excel.py        # Writes JSON data into an Excel template
├── gen_template.py       # Generates the blank rule template on Desktop
├── presets/              # Saved sub-schemes (created during usage)
│   └── DataTester-Game.md   # Example: game data classification preset
└── README.md
```

## Example preset: DataTester-Game

`presets/DataTester-Game.md` is an example sub-scheme for classifying game telemetry fields before sharing with a third-party game analytics company.

**Business rules encoded:**
- **明文传输** (plaintext): data whose primary purpose is game behavior analysis (combat actions, level progress, skill usage, AI behavior, etc.)
- **匿名化传输** (anonymized): data that is both game-related and personally identifiable (e.g. account ID, device ID)
- **不传输** (do not transmit): device/system metrics (memory, CPU, device performance scores) — regardless of whether they were collected during gameplay — plus pure personal identity data with no game-behavior purpose

To use it, load "DataTester-Game" when prompted at startup and provide your own input file path.

## Helper scripts

### `read_excel.py`

```bash
# Read rows 2–50 from a file (row 1 is header by default)
python3 ~/.claude/skills/DataTester/read_excel.py file.xlsx 2 50

# Get total row count
python3 ~/.claude/skills/DataTester/read_excel.py file.xlsx count

# Read from row 5, treating row 4 as the header (for files with title rows above the header)
python3 ~/.claude/skills/DataTester/read_excel.py file.xlsx 5 999 --header-row 4
```

Outputs JSON to stdout with `headers`, `rows`, `merge_info`, and metadata.

### `write_excel.py`

```bash
python3 ~/.claude/skills/DataTester/write_excel.py template.xlsx data.json output.xlsx
```

Copies the template (preserving styles and column widths), clears data rows, and writes new rows from the JSON file. The JSON format:

```json
{
  "headers": ["Col A", "Col B", ...],
  "rows": [
    {"Col A": "value", "Col B": "value", ...},
    ...
  ]
}
```

## License

MIT
