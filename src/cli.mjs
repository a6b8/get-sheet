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
  tabs                List all tabs
  addtab              Add a new tab
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

Options (addtab):
  --name <name>         Name for the new tab

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

    if( command === 'chart' ) {
        const { tab, range, title } = values
        const type = values[ 'type' ]
        const result = await GetSheetCli.chart( { tab, range, type, title, cwd } )
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
