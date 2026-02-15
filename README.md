[![Test](https://img.shields.io/github/actions/workflow/status/a6b8/get-sheet/test-on-release.yml)]()
[![Codecov](https://img.shields.io/codecov/c/github/a6b8/get-sheet)]()
![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)

```
   ______     __  _____ __              __
  / ____/__  / /_/ ___// /_  ___  ___  / /_
 / / __/ _ \/ __/\__ \/ __ \/ _ \/ _ \/ __/
/ /_/ /  __/ /_ ___/ / / / /  __/  __/ /_
\____/\___/\__//____/_/ /_/\___/\___/\__/
```

GetSheet is a Node.js CLI tool for reading, writing, formatting and managing Google Sheets from the terminal. It acts as a shared sketchpad between you and your AI agent — push tables, pull data, format cells, create charts.

## Quickstart

```bash
npm install -g get-sheet

getsheet init --credentials ~/keys/service-account.json --spreadsheet <SPREADSHEET_ID>
getsheet info
getsheet tabs
```

## Features

- **Read & write** — Push and pull data as 2D arrays
- **Append rows** — Add data without overwriting existing content
- **Formatting** — Bold, colors, font size, alignment, text wrapping, font family
- **Conditional formatting** — Color scales, value-based rules, custom formulas
- **Tab management** — List, create, color, and delete tabs
- **Layout control** — Freeze rows/cols, auto-filter, column widths, row heights, hide/unhide
- **Charts** — BAR, LINE, PIE, COLUMN, AREA, SCATTER
- **Two-tier config** — Global credentials in `~/.gsheet/`, per-project spreadsheet in `.gsheet/`
- **Service account auth** — Uses Google service account JSON key

## Table of Contents

- [Quickstart](#quickstart)
- [Features](#features)
- [Setup](#setup)
- [Commands](#commands)
- [Configuration](#configuration)
- [License](#license)

## Setup

```
Google Cloud Console → Enable Sheets API → Create Service Account → Download JSON Key → getsheet init → Share Sheet → Ready
```

1. Create a Google Cloud project and enable the **Google Sheets API**
2. Create a **service account** and download the JSON key
3. Run `getsheet init --credentials <path-to-json> --spreadsheet <id>`
4. Run `getsheet info` to see the service account email
5. Share your Google Sheet with that email (Editor role)

The Spreadsheet ID is in the URL:

```
https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
```

## Commands

| Command | Description |
|---------|-------------|
| `init` | Initialize credentials and link a spreadsheet |
| `read` | Read data from a tab |
| `write` | Write data to a tab |
| `append` | Append rows to a tab |
| `clear` | Clear a range |
| `delete` | Delete rows or columns from a tab |
| `format` | Format cells (bold, colors, alignment, font size, wrapping, font) |
| `condformat` | Add conditional formatting (color scale, conditions, formulas) |
| `clearcondformat` | Remove all conditional formatting from a tab |
| `freeze` | Freeze rows and/or columns |
| `filter` | Apply auto-filter to a range |
| `colwidth` | Set column width in pixels |
| `rowheight` | Set row height in pixels |
| `hide` | Hide rows or columns |
| `unhide` | Unhide rows or columns |
| `tabs` | List all tabs |
| `addtab` | Create a new tab |
| `tabcolor` | Set tab color |
| `chart` | Create a chart from a data range |
| `info` | Show service account email and setup info |

### Data Commands

```bash
# Read
getsheet read --tab Sheet1
getsheet read --tab Sheet1 --range A1:C10

# Write
getsheet write --tab Sheet1 --data '[["Name","Score"],["Alice",95],["Bob",87]]'
getsheet write --tab Sheet1 --range A1 --data '[["Name","Score"]]'

# Append
getsheet append --tab Sheet1 --data '[["Charlie",92]]'

# Clear
getsheet clear --tab Sheet1 --range B2:B10

# Delete rows or columns
getsheet delete --tab Sheet1 --rows 2:5
getsheet delete --tab Sheet1 --cols B:C
```

### Formatting Commands

```bash
# Bold, background, text color, font size, alignment, wrapping, font
getsheet format --tab Sheet1 --range A1:O1 --bold --bg "#f0f0f0" --color "#333333"
getsheet format --tab Sheet1 --range A1:O1 --fontsize 12
getsheet format --tab Sheet1 --range B2:O44 --align center
getsheet format --tab Sheet1 --range E7:E43 --wrap wrap
getsheet format --tab Sheet1 --range A1:O44 --font "Roboto Mono"

# Conditional formatting: color scale
getsheet condformat --tab Sheet1 --range B2:N44 --scale "red:yellow:green"

# Conditional formatting: value-based
getsheet condformat --tab Sheet1 --range O2:O44 --gt 100 --bg "#4caf50"
getsheet condformat --tab Sheet1 --range B2:N44 --between "8:10" --bg "#c8e6c9" --bold

# Conditional formatting: custom formula
getsheet condformat --tab Sheet1 --range A2:A44 --formula '=A2>100' --bg "#ffcdd2"

# Clear all conditional formatting
getsheet clearcondformat --tab Sheet1
```

### Layout Commands

```bash
# Freeze header row and first column
getsheet freeze --tab Sheet1 --rows 1 --cols 1

# Auto-filter
getsheet filter --tab Sheet1 --range A1:D100

# Column widths and row heights
getsheet colwidth --tab Sheet1 --cols A:C --width 150
getsheet rowheight --tab Sheet1 --rows 7:43 --height 21

# Hide / unhide
getsheet hide --tab Sheet1 --cols H:Z
getsheet unhide --tab Sheet1 --rows 2:5
```

### Tab & Chart Commands

```bash
# List tabs
getsheet tabs

# Create tab and set color
getsheet addtab --name Benchmarks
getsheet tabcolor --tab Benchmarks --color "#4285f4"

# Create chart
getsheet chart --tab Sheet1 --range A1:B4 --type BAR --title "Scores"
```

## Configuration

GetSheet uses a two-tier configuration:

```
~/.gsheet/config.json (Global: credentials path) ─┐
                                                   ├─→ getsheet read/write/...
.gsheet/config.json (Local: spreadsheet ID) ───────┘
```

### Global Config

Stored at `~/.gsheet/config.json`. Created on first `init`.

```json
{
    "credentials": "/path/to/service-account.json"
}
```

### Local Config

Stored at `.gsheet/config.json` in your working directory. Created per project via `init --spreadsheet`.

```json
{
    "root": "~/.gsheet",
    "spreadsheet": "1dXXCi8Dc5ZiMh8IvwUlN08jBz-yCFczF_qpQ3wU3Kdw"
}
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
