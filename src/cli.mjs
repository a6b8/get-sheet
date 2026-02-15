#!/usr/bin/env node
import { parseArgs } from 'node:util'

import { GetSheetCli } from './GetSheetCli.mjs'


const args = parseArgs( {
    args: process.argv.slice( 2 ),
    allowPositionals: true,
    strict: false,
    options: {
        'credentials': { type: 'string' },
        'spreadsheet': { type: 'string' },
        'tab': { type: 'string' },
        'range': { type: 'string' },
        'data': { type: 'string' },
        'name': { type: 'string' },
        'type': { type: 'string' },
        'title': { type: 'string' },
        'rows': { type: 'string' },
        'cols': { type: 'string' },
        'bold': { type: 'boolean' },
        'bg': { type: 'string' },
        'color': { type: 'string' },
        'fontsize': { type: 'string' },
        'align': { type: 'string' },
        'wrap': { type: 'string' },
        'font': { type: 'string' },
        'scale': { type: 'string' },
        'min': { type: 'string' },
        'max': { type: 'string' },
        'gt': { type: 'string' },
        'lt': { type: 'string' },
        'eq': { type: 'string' },
        'between': { type: 'string' },
        'formula': { type: 'string' },
        'width': { type: 'string' },
        'height': { type: 'string' },
        'help': { type: 'boolean', short: 'h' }
    }
} )

const { positionals, values } = args
const command = positionals[ 0 ]
const cwd = process.cwd()

const output = ( { result } ) => {
    process.stdout.write( JSON.stringify( result, null, 4 ) + '\n' )
}

const showHelp = () => {
    const helpText = `
   ______     __  _____ __              __
  / ____/__  / /_/ ___// /_  ___  ___  / /_
 / / __/ _ \\/ __/\\__ \\/ __ \\/ _ \\/ _ \\/ __/
/ /_/ /  __/ /_ ___/ / / / /  __/  __/ /_
\\____/\\___/\\__//____/_/ /_/\\___/\\___/\\__/

Usage: getsheet <command> [options]

Commands:
  init                Initialize credentials and spreadsheet
  read                Read data from spreadsheet
  write               Write data to spreadsheet
  append              Append rows to spreadsheet
  clear               Clear a range in spreadsheet
  delete              Delete rows or columns from a tab
  format              Format cells (bold, colors, alignment, font size)
  condformat          Add conditional formatting (color scale or rules)
  clearcondformat     Remove all conditional formatting from a tab
  freeze              Freeze rows and/or columns
  filter              Apply auto-filter to a range
  colwidth            Set column width in pixels
  rowheight           Set row height in pixels
  hide                Hide rows or columns
  unhide              Unhide rows or columns
  tabs                List all tabs
  addtab              Add a new tab
  tabcolor            Set tab color
  chart               Create a chart from data range
  info                Show service account email and setup info

Options (init):
  --credentials <path>  Path to Google service account JSON
  --spreadsheet <id>    Google Spreadsheet ID

Options (read):
  --tab <name>          Tab name (default: Sheet1)
  --range <range>       Cell range, e.g. A1:D10 (default: all)

Options (write):
  --tab <name>          Tab name (required)
  --range <range>       Start range, e.g. A1 (default: start of tab)
  --data '<json>'       2D array as JSON string (required)

Options (append):
  --tab <name>          Tab name (required)
  --data '<json>'       2D array as JSON string (required)

Options (clear):
  --tab <name>          Tab name (default: Sheet1)
  --range <range>       Cell range to clear (default: all)

Options (delete):
  --tab <name>          Tab name (required)
  --rows <range>        Row range to delete, e.g. "2:5" (1-based, inclusive)
  --cols <range>        Column range to delete, e.g. "B:C" (letter-based, inclusive)
                        One of --rows or --cols is required

Options (format):
  --tab <name>          Tab name (required)
  --range <range>       Cell range, e.g. A1:O1 (required)
  --bold                Make text bold
  --bg <hex>            Background color, e.g. "#4285f4"
  --color <hex>         Text color, e.g. "#333333"
  --fontsize <n>        Font size, e.g. 12
  --align <align>       Horizontal alignment: left, center, right
  --wrap <mode>         Text wrapping: wrap, clip, overflow
  --font <name>         Font family, e.g. "Roboto Mono"
                        At least one format option is required

Options (condformat):
  --tab <name>          Tab name (required)
  --range <range>       Cell range, e.g. B2:N44 (required)
  Color scale mode:
    --scale <colors>    2-3 colors separated by ":", e.g. "red:yellow:green"
                        Named: red, green, yellow, white, orange, blue
                        Hex: "#ff0000:#ffff00:#00ff00"
    --min <n>           Min value for scale (optional, default: auto)
    --max <n>           Max value for scale (optional, default: auto)
  Condition mode:
    --gt <n>            Greater than
    --lt <n>            Less than
    --eq <n>            Equal to
    --between <range>   Between range, e.g. "8:10"
    --bg <hex>          Background color for matching cells (required)
    --bold              Bold text for matching cells (optional)
  Formula mode:
    --formula <expr>    Custom formula, e.g. "=A1>100"
    --bg <hex>          Background color for matching cells (required)
    --bold              Bold text for matching cells (optional)

Options (clearcondformat):
  --tab <name>          Tab name (required)

Options (freeze):
  --tab <name>          Tab name (required)
  --rows <n>            Number of rows to freeze (e.g. 1 for header)
  --cols <n>            Number of columns to freeze (e.g. 1 for first col)
                        At least one of --rows or --cols is required
                        Use 0 to unfreeze: --rows 0

Options (filter):
  --tab <name>          Tab name (required)
  --range <range>       Cell range to filter, e.g. A1:D100 (required)

Options (colwidth):
  --tab <name>          Tab name (required)
  --cols <range>        Column or column range, e.g. "A" or "A:C" (required)
  --width <pixels>      Width in pixels, e.g. 150 (required)

Options (rowheight):
  --tab <name>          Tab name (required)
  --rows <range>        Row or row range, e.g. "7" or "7:43" (1-based, inclusive)
  --height <pixels>     Height in pixels, e.g. 21 (required)

Options (hide / unhide):
  --tab <name>          Tab name (required)
  --rows <range>        Row range, e.g. "2:5" (1-based, inclusive)
  --cols <range>        Column range, e.g. "B:C" (letter-based, inclusive)
                        One of --rows or --cols is required

Options (addtab):
  --name <name>         Name for the new tab

Options (tabcolor):
  --tab <name>          Tab name (required)
  --color <hex>         Tab color, e.g. "#4285f4" (required)

Options (chart):
  --tab <name>          Tab with the data (required)
  --range <range>       Data range, e.g. A1:B5 (required)
  --type <type>         Chart type: BAR, LINE, PIE, COLUMN, AREA, SCATTER
                        (default: COLUMN)
  --title <title>       Chart title (optional)

General:
  --help, -h            Show this help message

Setup:
  1. Create a Google Cloud project and enable Sheets API
  2. Create a service account and download the JSON key
  3. Run: getsheet init --credentials <path-to-json> --spreadsheet <sheet-id>
  4. Share your spreadsheet with the service account email (Editor role)
  5. Run: getsheet info    to see the email address to share with

The Spreadsheet ID is in the URL:
  https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit

Examples:
  getsheet init --credentials ~/keys/sa.json --spreadsheet 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms
  getsheet info
  getsheet tabs
  getsheet read --tab Sheet1
  getsheet read --tab Sheet1 --range A1:C10
  getsheet write --tab Sheet1 --data '[["Name","Score"],["Alice",95],["Bob",87]]'
  getsheet write --tab Sheet1 --range A1 --data '[["Name","Score"]]'
  getsheet append --tab Sheet1 --data '[["Charlie",92]]'
  getsheet clear --tab Sheet1 --range B2:B10
  getsheet addtab --name Benchmarks
  getsheet delete --tab Sheet1 --rows 2:5
  getsheet delete --tab Sheet1 --cols B:C
  getsheet format --tab Sheet1 --range A1:O1 --bold
  getsheet format --tab Sheet1 --range A1:A44 --bg "#4285f4"
  getsheet format --tab Sheet1 --range A1:O1 --bold --bg "#f0f0f0" --color "#333333"
  getsheet format --tab Sheet1 --range B2:O44 --align center
  getsheet format --tab Sheet1 --range A1:O1 --fontsize 12
  getsheet format --tab Sheet1 --range A1:O44 --font "Roboto Mono"
  getsheet format --tab Sheet1 --range E7:E43 --wrap wrap
  getsheet condformat --tab Sheet1 --range B2:N44 --scale "red:yellow:green" --min 0 --max 10
  getsheet condformat --tab Sheet1 --range B2:N44 --scale "red:green"
  getsheet condformat --tab Sheet1 --range O2:O44 --gt 100 --bg "#4caf50"
  getsheet condformat --tab Sheet1 --range B2:N44 --between "8:10" --bg "#c8e6c9" --bold
  getsheet condformat --tab Sheet1 --range A2:A44 --formula "=A2>100" --bg "#ffcdd2"
  getsheet clearcondformat --tab Sheet1
  getsheet tabcolor --tab Sheet1 --color "#4285f4"
  getsheet freeze --tab Sheet1 --rows 1
  getsheet freeze --tab Sheet1 --rows 1 --cols 1
  getsheet freeze --tab Sheet1 --rows 0 --cols 0
  getsheet filter --tab Sheet1 --range A1:D100
  getsheet colwidth --tab Sheet1 --cols A --width 200
  getsheet colwidth --tab Sheet1 --cols A:C --width 150
  getsheet rowheight --tab Sheet1 --rows 7:43 --height 21
  getsheet rowheight --tab Sheet1 --rows 1 --height 30
  getsheet hide --tab Sheet1 --rows 2:5
  getsheet hide --tab Sheet1 --cols B:C
  getsheet unhide --tab Sheet1 --rows 2:5
  getsheet unhide --tab Sheet1 --cols B:C
  getsheet chart --tab Sheet1 --range A1:B4 --type BAR --title "Scores"
`

    process.stdout.write( helpText )
}

const run = async () => {
    if( values[ 'help' ] || !command ) {
        showHelp()

        return
    }

    if( command === 'init' ) {
        const { credentials, spreadsheet } = values
        const result = await GetSheetCli.init( { credentials, spreadsheet, cwd } )
        output( { result } )

        return
    }

    if( command === 'read' ) {
        const { tab, range } = values
        const result = await GetSheetCli.read( { tab, range, cwd } )
        output( { result } )

        return
    }

    if( command === 'write' ) {
        const { tab, range } = values
        const rawData = values[ 'data' ]
        let data

        try {
            data = JSON.parse( rawData )
        } catch {
            output( { result: { 'status': false, 'error': '--data must be valid JSON. Example: [["a","b"],["c","d"]]' } } )

            return
        }

        const result = await GetSheetCli.write( { tab, range, data, cwd } )
        output( { result } )

        return
    }

    if( command === 'append' ) {
        const { tab } = values
        const rawData = values[ 'data' ]
        let data

        try {
            data = JSON.parse( rawData )
        } catch {
            output( { result: { 'status': false, 'error': '--data must be valid JSON. Example: [["a","b"],["c","d"]]' } } )

            return
        }

        const result = await GetSheetCli.append( { tab, data, cwd } )
        output( { result } )

        return
    }

    if( command === 'clear' ) {
        const { tab, range } = values
        const result = await GetSheetCli.clear( { tab, range, cwd } )
        output( { result } )

        return
    }

    if( command === 'tabs' ) {
        const result = await GetSheetCli.tabs( { cwd } )
        output( { result } )

        return
    }

    if( command === 'addtab' ) {
        const { name } = values
        const result = await GetSheetCli.addTab( { name, cwd } )
        output( { result } )

        return
    }

    if( command === 'tabcolor' ) {
        const { tab, color } = values
        const result = await GetSheetCli.tabColor( { tab, color, cwd } )
        output( { result } )

        return
    }

    if( command === 'chart' ) {
        const { tab, range, title } = values
        const type = values[ 'type' ]
        const result = await GetSheetCli.chart( { tab, range, type, title, cwd } )
        output( { result } )

        return
    }

    if( command === 'delete' ) {
        const { tab, rows, cols } = values
        const result = await GetSheetCli.delete( { tab, rows, cols, cwd } )
        output( { result } )

        return
    }

    if( command === 'format' ) {
        const { tab, range, bg, color, fontsize, align, wrap, font } = values
        const bold = values[ 'bold' ]
        const result = await GetSheetCli.format( { tab, range, bold, bg, color, fontsize, align, wrap, font, cwd } )
        output( { result } )

        return
    }

    if( command === 'condformat' ) {
        const { tab, range, bg, between, formula } = values
        const scale = values[ 'scale' ]
        const min = values[ 'min' ]
        const max = values[ 'max' ]
        const gt = values[ 'gt' ]
        const lt = values[ 'lt' ]
        const eq = values[ 'eq' ]
        const bold = values[ 'bold' ]
        const result = await GetSheetCli.condFormat( { tab, range, scale, min, max, gt, lt, eq, between, formula, bg, bold, cwd } )
        output( { result } )

        return
    }

    if( command === 'clearcondformat' ) {
        const { tab } = values
        const result = await GetSheetCli.clearCondFormat( { tab, cwd } )
        output( { result } )

        return
    }

    if( command === 'hide' ) {
        const { tab, rows, cols } = values
        const result = await GetSheetCli.hide( { tab, rows, cols, cwd } )
        output( { result } )

        return
    }

    if( command === 'unhide' ) {
        const { tab, rows, cols } = values
        const result = await GetSheetCli.unhide( { tab, rows, cols, cwd } )
        output( { result } )

        return
    }

    if( command === 'freeze' ) {
        const { tab, rows, cols } = values
        const result = await GetSheetCli.freeze( { tab, rows, cols, cwd } )
        output( { result } )

        return
    }

    if( command === 'filter' ) {
        const { tab, range } = values
        const result = await GetSheetCli.filter( { tab, range, cwd } )
        output( { result } )

        return
    }

    if( command === 'colwidth' ) {
        const { tab, cols, width } = values
        const result = await GetSheetCli.colWidth( { tab, cols, width, cwd } )
        output( { result } )

        return
    }

    if( command === 'rowheight' ) {
        const { tab, rows, height } = values
        const result = await GetSheetCli.rowHeight( { tab, rows, height, cwd } )
        output( { result } )

        return
    }

    if( command === 'info' ) {
        const result = await GetSheetCli.info( { cwd } )
        output( { result } )

        return
    }

    showHelp()
}

run()
