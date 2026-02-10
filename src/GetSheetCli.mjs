import { readFile, writeFile, mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { homedir } from 'node:os'

import { google } from 'googleapis'


class GetSheetCli {
    static async init( { credentials, spreadsheet, cwd } ) {
        const globalDir = GetSheetCli.#gsheetDir()
        const globalConfigPath = join( globalDir, 'config.json' )
        const { data: existingGlobalConfig } = await GetSheetCli.#readJson( { filePath: globalConfigPath } )

        const credPath = credentials || ( existingGlobalConfig ? existingGlobalConfig['credentials'] : undefined )

        const { status, error } = GetSheetCli.validationInit( { credentials: credPath, spreadsheet } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
        }

        const { data: credJson } = await GetSheetCli.#readJson( { filePath: credPath } )
        if( !credJson ) {
            return GetSheetCli.#error( { error: `Could not read credentials file: ${credPath}` } )
        }

        const { client_email: clientEmail } = credJson
        if( !clientEmail ) {
            return GetSheetCli.#error( { error: 'Credentials file missing "client_email" field' } )
        }

        if( credentials ) {
            await mkdir( globalDir, { recursive: true } )

            const globalConfig = {
                'credentials': credentials
            }
            await writeFile( globalConfigPath, JSON.stringify( globalConfig, null, 4 ), 'utf-8' )
        }

        const localConfigDir = join( cwd, '.gsheet' )
        await mkdir( localConfigDir, { recursive: true } )

        const localConfigPath = join( localConfigDir, 'config.json' )
        const localConfig = {
            'root': '~/.gsheet',
            'spreadsheet': spreadsheet
        }
        await writeFile( localConfigPath, JSON.stringify( localConfig, null, 4 ), 'utf-8' )

        const result = {
            'status': true,
            'message': `Initialized. Share your spreadsheet with: ${clientEmail}`,
            'clientEmail': clientEmail,
            'spreadsheet': spreadsheet
        }

        return result
    }


    static async read( { tab, range, cwd } ) {
        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config
        const sheetRange = range
            ? `${tab || 'Sheet1'}!${range}`
            : tab || 'Sheet1'

        try {
            const response = await sheets.spreadsheets.values.get( {
                'spreadsheetId': spreadsheet,
                'range': sheetRange
            } )

            const { data } = response
            const rows = data['values'] || []

            const result = {
                'status': true,
                'range': data['range'],
                'rows': rows.length,
                'data': rows
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Read failed: ${err.message}` } )
        }
    }


    static async write( { tab, range, data, cwd } ) {
        const { status, error } = GetSheetCli.validationWrite( { tab, data } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
        }

        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config
        const sheetRange = range
            ? `${tab}!${range}`
            : tab

        try {
            const response = await sheets.spreadsheets.values.update( {
                'spreadsheetId': spreadsheet,
                'range': sheetRange,
                'valueInputOption': 'USER_ENTERED',
                'requestBody': {
                    'values': data
                }
            } )

            const { data: responseData } = response
            const result = {
                'status': true,
                'updatedRange': responseData['updatedRange'],
                'updatedRows': responseData['updatedRows'],
                'updatedColumns': responseData['updatedColumns'],
                'updatedCells': responseData['updatedCells']
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Write failed: ${err.message}` } )
        }
    }


    static async append( { tab, data, cwd } ) {
        const { status, error } = GetSheetCli.validationWrite( { tab, data } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
        }

        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config

        try {
            const response = await sheets.spreadsheets.values.append( {
                'spreadsheetId': spreadsheet,
                'range': tab,
                'valueInputOption': 'USER_ENTERED',
                'requestBody': {
                    'values': data
                }
            } )

            const { data: responseData } = response
            const { updates } = responseData
            const result = {
                'status': true,
                'updatedRange': updates['updatedRange'],
                'updatedRows': updates['updatedRows'],
                'updatedColumns': updates['updatedColumns'],
                'updatedCells': updates['updatedCells']
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Append failed: ${err.message}` } )
        }
    }


    static async clear( { tab, range, cwd } ) {
        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config
        const sheetRange = range
            ? `${tab || 'Sheet1'}!${range}`
            : tab || 'Sheet1'

        try {
            await sheets.spreadsheets.values.clear( {
                'spreadsheetId': spreadsheet,
                'range': sheetRange
            } )

            const result = {
                'status': true,
                'clearedRange': sheetRange
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Clear failed: ${err.message}` } )
        }
    }


    static async tabs( { cwd } ) {
        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config

        try {
            const response = await sheets.spreadsheets.get( {
                'spreadsheetId': spreadsheet,
                'fields': 'sheets.properties'
            } )

            const { data } = response
            const tabList = data['sheets']
                .map( ( sheet ) => {
                    const { properties } = sheet
                    const entry = {
                        'title': properties['title'],
                        'index': properties['index'],
                        'sheetId': properties['sheetId'],
                        'rowCount': properties['gridProperties']['rowCount'],
                        'columnCount': properties['gridProperties']['columnCount']
                    }

                    return entry
                } )

            const result = {
                'status': true,
                'total': tabList.length,
                'tabs': tabList
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Tabs failed: ${err.message}` } )
        }
    }


    static async addTab( { name, cwd } ) {
        if( name === undefined ) {
            return GetSheetCli.#error( { error: '--name is required. Provide a tab name' } )
        }

        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config

        try {
            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [
                        {
                            'addSheet': {
                                'properties': {
                                    'title': name
                                }
                            }
                        }
                    ]
                }
            } )

            const result = {
                'status': true,
                'message': `Tab "${name}" created`,
                'name': name
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Add tab failed: ${err.message}` } )
        }
    }


    static async chart( { tab, range, type, title, cwd } ) {
        if( tab === undefined ) {
            return GetSheetCli.#error( { error: '--tab is required. Provide tab name' } )
        }

        if( range === undefined ) {
            return GetSheetCli.#error( { error: '--range is required. e.g. A1:B10' } )
        }

        const chartTypes = [ 'BAR', 'LINE', 'PIE', 'COLUMN', 'AREA', 'SCATTER' ]
        const chartType = ( type || 'COLUMN' ).toUpperCase()

        if( !chartTypes.includes( chartType ) ) {
            return GetSheetCli.#error( { error: `--type must be one of: ${chartTypes.join( ', ' )}` } )
        }

        const { config, error: configError } = await GetSheetCli.#loadConfig( { cwd } )
        if( !config ) {
            return GetSheetCli.#error( { error: configError } )
        }

        const { sheets, spreadsheet } = config

        try {
            const metaResponse = await sheets.spreadsheets.get( {
                'spreadsheetId': spreadsheet,
                'fields': 'sheets.properties'
            } )

            const { data: metaData } = metaResponse
            const sheetMeta = metaData['sheets']
                .find( ( s ) => {
                    const matches = s['properties']['title'] === tab

                    return matches
                } )

            if( !sheetMeta ) {
                return GetSheetCli.#error( { error: `Tab "${tab}" not found` } )
            }

            const sheetId = sheetMeta['properties']['sheetId']
            const { startCol, startRow, endCol, endRow } = GetSheetCli.#parseRange( { range } )

            const chartRequest = {
                'addChart': {
                    'chart': {
                        'spec': {
                            'title': title || '',
                            'basicChart': {
                                'chartType': chartType,
                                'legendPosition': 'BOTTOM_LEGEND',
                                'domains': [
                                    {
                                        'domain': {
                                            'sourceRange': {
                                                'sources': [
                                                    {
                                                        'sheetId': sheetId,
                                                        'startRowIndex': startRow,
                                                        'endRowIndex': endRow,
                                                        'startColumnIndex': startCol,
                                                        'endColumnIndex': startCol + 1
                                                    }
                                                ]
                                            }
                                        }
                                    }
                                ],
                                'series': GetSheetCli.#buildChartSeries( { sheetId, startRow, endRow, startCol: startCol + 1, endCol } ),
                                'headerCount': 1
                            }
                        },
                        'position': {
                            'overlayPosition': {
                                'anchorCell': {
                                    'sheetId': sheetId,
                                    'rowIndex': endRow + 1,
                                    'columnIndex': startCol
                                }
                            }
                        }
                    }
                }
            }

            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [ chartRequest ]
                }
            } )

            const result = {
                'status': true,
                'message': `${chartType} chart created from ${tab}!${range}`,
                'chartType': chartType,
                'dataRange': `${tab}!${range}`,
                'title': title || '(untitled)'
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Chart failed: ${err.message}` } )
        }
    }


    static async info( { cwd } ) {
        const globalConfigPath = join( GetSheetCli.#gsheetDir(), 'config.json' )
        const { data: globalConfig } = await GetSheetCli.#readJson( { filePath: globalConfigPath } )

        if( !globalConfig ) {
            return GetSheetCli.#error( { error: 'Not initialized. Run: getsheet init --credentials <path> --spreadsheet <id>' } )
        }

        const { credentials: credPath } = globalConfig
        const { data: credJson } = await GetSheetCli.#readJson( { filePath: credPath } )

        if( !credJson ) {
            return GetSheetCli.#error( { error: `Could not read credentials file: ${credPath}` } )
        }

        const { client_email: clientEmail, project_id: projectId } = credJson

        const localConfigPath = join( cwd, '.gsheet', 'config.json' )
        const { data: localConfig } = await GetSheetCli.#readJson( { filePath: localConfigPath } )

        const result = {
            'status': true,
            'clientEmail': clientEmail,
            'projectId': projectId || null,
            'spreadsheet': localConfig ? localConfig['spreadsheet'] : null,
            'shareInstructions': `Share your Google Sheet with: ${clientEmail} (Editor role)`
        }

        return result
    }


    static validationInit( { credentials, spreadsheet } ) {
        const struct = { 'status': false, 'error': null }

        if( credentials === undefined ) {
            struct['error'] = '--credentials is required (first-time setup). Provide path to Google service account JSON'

            return struct
        }

        if( spreadsheet === undefined ) {
            struct['error'] = '--spreadsheet is required. Provide Google Spreadsheet ID'

            return struct
        }

        struct['status'] = true

        return struct
    }


    static validationWrite( { tab, data } ) {
        const struct = { 'status': false, 'error': null }

        if( tab === undefined ) {
            struct['error'] = '--tab is required. Provide tab name'

            return struct
        }

        if( data === undefined ) {
            struct['error'] = '--data is required. Provide 2D array as JSON'

            return struct
        }

        if( !Array.isArray( data ) ) {
            struct['error'] = '--data must be a 2D array. Example: [["a","b"],["c","d"]]'

            return struct
        }

        if( data.length === 0 ) {
            struct['error'] = '--data must not be empty'

            return struct
        }

        const hasInvalidRow = data
            .some( ( row ) => {
                const invalid = !Array.isArray( row )

                return invalid
            } )

        if( hasInvalidRow ) {
            struct['error'] = '--data must be a 2D array. Each row must be an array'

            return struct
        }

        struct['status'] = true

        return struct
    }


    static async #loadConfig( { cwd } ) {
        const globalConfigPath = join( GetSheetCli.#gsheetDir(), 'config.json' )
        const { data: globalConfig } = await GetSheetCli.#readJson( { filePath: globalConfigPath } )

        if( !globalConfig ) {
            return { 'config': null, 'error': 'Not initialized. Run: getsheet init --credentials <path> --spreadsheet <id>' }
        }

        const localConfigPath = join( cwd, '.gsheet', 'config.json' )
        const { data: localConfig } = await GetSheetCli.#readJson( { filePath: localConfigPath } )

        if( !localConfig ) {
            return { 'config': null, 'error': 'No local config. Run: getsheet init --spreadsheet <id>' }
        }

        const { credentials: credPath } = globalConfig
        const { spreadsheet } = localConfig

        const { auth, error: authError } = await GetSheetCli.#getAuth( { credentials: credPath } )
        if( !auth ) {
            return { 'config': null, 'error': authError }
        }

        const { sheets } = GetSheetCli.#getSheets( { auth } )

        const config = {
            'credentials': credPath,
            spreadsheet,
            sheets,
            'local': localConfig
        }

        return { config, 'error': null }
    }


    static async #getAuth( { credentials } ) {
        try {
            const auth = new google.auth.GoogleAuth( {
                'keyFile': credentials,
                'scopes': [ 'https://www.googleapis.com/auth/spreadsheets' ]
            } )

            return { auth, 'error': null }
        } catch( err ) {
            return { 'auth': null, 'error': `Auth failed: ${err.message}` }
        }
    }


    static #getSheets( { auth } ) {
        const sheets = google.sheets( { 'version': 'v4', auth } )

        return { sheets }
    }


    static #gsheetDir() {
        const dir = join( homedir(), '.gsheet' )

        return dir
    }


    static async #readJson( { filePath } ) {
        try {
            const content = await readFile( filePath, 'utf-8' )
            const data = JSON.parse( content )

            return { data }
        } catch {
            return { 'data': null }
        }
    }


    static #parseRange( { range } ) {
        const match = range.match( /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/ )
        if( !match ) {
            return { 'startCol': 0, 'startRow': 0, 'endCol': 1, 'endRow': 1 }
        }

        const colToIndex = ( col ) => {
            const index = col
                .split( '' )
                .reduce( ( acc, char ) => {
                    const val = acc * 26 + char.charCodeAt( 0 ) - 64

                    return val
                }, 0 ) - 1

            return index
        }

        const startCol = colToIndex( match[ 1 ] )
        const startRow = parseInt( match[ 2 ] ) - 1
        const endCol = colToIndex( match[ 3 ] ) + 1
        const endRow = parseInt( match[ 4 ] )

        return { startCol, startRow, endCol, endRow }
    }


    static #buildChartSeries( { sheetId, startRow, endRow, startCol, endCol } ) {
        const series = []

        Array.from( { length: endCol - startCol }, ( _, i ) => {
            const colIndex = startCol + i
            series.push( {
                'series': {
                    'sourceRange': {
                        'sources': [
                            {
                                sheetId,
                                'startRowIndex': startRow,
                                'endRowIndex': endRow,
                                'startColumnIndex': colIndex,
                                'endColumnIndex': colIndex + 1
                            }
                        ]
                    }
                }
            } )
        } )

        return series
    }


    static #error( { error } ) {
        const result = { 'status': false, error }

        return result
    }
}


export { GetSheetCli }
