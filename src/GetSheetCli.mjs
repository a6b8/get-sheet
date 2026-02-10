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


    static async delete( { tab, rows, cols, cwd } ) {
        const { status, error } = GetSheetCli.validationDelete( { tab, rows, cols } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
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

            let dimension
            let startIndex
            let endIndex

            if( rows ) {
                const parts = rows.split( ':' )
                dimension = 'ROWS'
                startIndex = parseInt( parts[ 0 ] ) - 1
                endIndex = parseInt( parts[ 1 ] )
            } else {
                const colToIndex = ( col ) => {
                    const index = col.toUpperCase()
                        .split( '' )
                        .reduce( ( acc, char ) => {
                            const val = acc * 26 + char.charCodeAt( 0 ) - 64

                            return val
                        }, 0 ) - 1

                    return index
                }

                const parts = cols.split( ':' )
                dimension = 'COLUMNS'
                startIndex = colToIndex( parts[ 0 ] )
                endIndex = colToIndex( parts[ 1 ] ) + 1
            }

            const deleteRequest = {
                'deleteDimension': {
                    'range': {
                        'sheetId': sheetId,
                        'dimension': dimension,
                        'startIndex': startIndex,
                        'endIndex': endIndex
                    }
                }
            }

            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [ deleteRequest ]
                }
            } )

            const rangeLabel = rows ? `rows ${rows}` : `columns ${cols}`
            const result = {
                'status': true,
                'message': `Deleted ${rangeLabel} from "${tab}"`,
                'tab': tab,
                'dimension': dimension,
                'startIndex': startIndex,
                'endIndex': endIndex
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Delete failed: ${err.message}` } )
        }
    }


    static async format( { tab, range, bold, bg, color, fontsize, align, cwd } ) {
        const { status, error } = GetSheetCli.validationFormat( { tab, range, bold, bg, color, fontsize, align } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
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

            const userEnteredFormat = {}
            const fieldParts = []

            if( bold ) {
                if( !userEnteredFormat['textFormat'] ) {
                    userEnteredFormat['textFormat'] = {}
                }
                userEnteredFormat['textFormat']['bold'] = true
                fieldParts.push( 'userEnteredFormat.textFormat.bold' )
            }

            if( bg ) {
                const { red, green, blue } = GetSheetCli.#hexToRgb( { hex: bg } )
                userEnteredFormat['backgroundColor'] = { red, green, blue }
                fieldParts.push( 'userEnteredFormat.backgroundColor' )
            }

            if( color ) {
                if( !userEnteredFormat['textFormat'] ) {
                    userEnteredFormat['textFormat'] = {}
                }
                const { red, green, blue } = GetSheetCli.#hexToRgb( { hex: color } )
                userEnteredFormat['textFormat']['foregroundColor'] = { red, green, blue }
                fieldParts.push( 'userEnteredFormat.textFormat.foregroundColor' )
            }

            if( fontsize ) {
                const size = parseInt( fontsize )
                if( !userEnteredFormat['textFormat'] ) {
                    userEnteredFormat['textFormat'] = {}
                }
                userEnteredFormat['textFormat']['fontSize'] = size
                fieldParts.push( 'userEnteredFormat.textFormat.fontSize' )
            }

            if( align ) {
                userEnteredFormat['horizontalAlignment'] = align.toUpperCase()
                fieldParts.push( 'userEnteredFormat.horizontalAlignment' )
            }

            const formatRequest = {
                'repeatCell': {
                    'range': {
                        'sheetId': sheetId,
                        'startRowIndex': startRow,
                        'endRowIndex': endRow,
                        'startColumnIndex': startCol,
                        'endColumnIndex': endCol
                    },
                    'cell': {
                        'userEnteredFormat': userEnteredFormat
                    },
                    'fields': fieldParts.join( ',' )
                }
            }

            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [ formatRequest ]
                }
            } )

            const appliedFormats = []
            if( bold ) { appliedFormats.push( 'bold' ) }
            if( bg ) { appliedFormats.push( `bg:${bg}` ) }
            if( color ) { appliedFormats.push( `color:${color}` ) }
            if( fontsize ) { appliedFormats.push( `fontsize:${fontsize}` ) }
            if( align ) { appliedFormats.push( `align:${align}` ) }

            const result = {
                'status': true,
                'message': `Formatted ${tab}!${range} with ${appliedFormats.join( ', ' )}`,
                'tab': tab,
                'range': range,
                'formats': appliedFormats
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Format failed: ${err.message}` } )
        }
    }


    static async condFormat( { tab, range, scale, min, max, gt, lt, eq, between, bg, bold, cwd } ) {
        const { status, error } = GetSheetCli.validationCondFormat( { tab, range, scale, gt, lt, eq, between, bg } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
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

            const rangeObj = {
                'sheetId': sheetId,
                'startRowIndex': startRow,
                'endRowIndex': endRow,
                'startColumnIndex': startCol,
                'endColumnIndex': endCol
            }

            let rule

            if( scale ) {
                const { colors } = GetSheetCli.#parseScale( { scale } )
                const gradientRule = {}

                if( min !== undefined ) {
                    gradientRule['minpoint'] = {
                        'color': colors[ 0 ],
                        'type': 'NUMBER',
                        'value': String( min )
                    }
                } else {
                    gradientRule['minpoint'] = {
                        'color': colors[ 0 ],
                        'type': 'MIN'
                    }
                }

                if( colors.length === 3 ) {
                    if( min !== undefined && max !== undefined ) {
                        const mid = ( parseFloat( min ) + parseFloat( max ) ) / 2
                        gradientRule['midpoint'] = {
                            'color': colors[ 1 ],
                            'type': 'NUMBER',
                            'value': String( mid )
                        }
                    } else {
                        gradientRule['midpoint'] = {
                            'color': colors[ 1 ],
                            'type': 'PERCENTILE',
                            'value': '50'
                        }
                    }
                }

                if( max !== undefined ) {
                    gradientRule['maxpoint'] = {
                        'color': colors[ colors.length - 1 ],
                        'type': 'NUMBER',
                        'value': String( max )
                    }
                } else {
                    gradientRule['maxpoint'] = {
                        'color': colors[ colors.length - 1 ],
                        'type': 'MAX'
                    }
                }

                rule = {
                    'ranges': [ rangeObj ],
                    'gradientRule': gradientRule
                }
            } else {
                let condition

                if( gt !== undefined ) {
                    condition = {
                        'type': 'NUMBER_GREATER',
                        'values': [ { 'userEnteredValue': String( gt ) } ]
                    }
                } else if( lt !== undefined ) {
                    condition = {
                        'type': 'NUMBER_LESS',
                        'values': [ { 'userEnteredValue': String( lt ) } ]
                    }
                } else if( eq !== undefined ) {
                    condition = {
                        'type': 'NUMBER_EQ',
                        'values': [ { 'userEnteredValue': String( eq ) } ]
                    }
                } else if( between !== undefined ) {
                    const parts = between.split( ':' )
                    condition = {
                        'type': 'NUMBER_BETWEEN',
                        'values': [
                            { 'userEnteredValue': parts[ 0 ] },
                            { 'userEnteredValue': parts[ 1 ] }
                        ]
                    }
                }

                const cellFormat = {}

                if( bg ) {
                    const { red, green, blue } = GetSheetCli.#hexToRgb( { hex: bg } )
                    cellFormat['backgroundColor'] = { red, green, blue }
                }

                if( bold ) {
                    cellFormat['textFormat'] = { 'bold': true }
                }

                rule = {
                    'ranges': [ rangeObj ],
                    'booleanRule': {
                        'condition': condition,
                        'format': cellFormat
                    }
                }
            }

            const request = {
                'addConditionalFormatRule': {
                    'rule': rule,
                    'index': 0
                }
            }

            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [ request ]
                }
            } )

            const ruleType = scale ? `color scale (${scale})` : 'conditional rule'
            const result = {
                'status': true,
                'message': `Added ${ruleType} to ${tab}!${range}`,
                'tab': tab,
                'range': range,
                'ruleType': scale ? 'gradient' : 'boolean'
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Conditional format failed: ${err.message}` } )
        }
    }


    static async freeze( { tab, rows, cols, cwd } ) {
        const { status, error } = GetSheetCli.validationFreeze( { tab, rows, cols } )
        if( !status ) {
            return GetSheetCli.#error( { error } )
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
            const gridProperties = {}
            const fields = []

            if( rows !== undefined ) {
                gridProperties['frozenRowCount'] = parseInt( rows )
                fields.push( 'gridProperties.frozenRowCount' )
            }

            if( cols !== undefined ) {
                gridProperties['frozenColumnCount'] = parseInt( cols )
                fields.push( 'gridProperties.frozenColumnCount' )
            }

            const freezeRequest = {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheetId,
                        'gridProperties': gridProperties
                    },
                    'fields': fields.join( ',' )
                }
            }

            await sheets.spreadsheets.batchUpdate( {
                'spreadsheetId': spreadsheet,
                'requestBody': {
                    'requests': [ freezeRequest ]
                }
            } )

            const parts = []
            if( rows !== undefined ) { parts.push( `${rows} row(s)` ) }
            if( cols !== undefined ) { parts.push( `${cols} column(s)` ) }

            const result = {
                'status': true,
                'message': `Froze ${parts.join( ' and ' )} in "${tab}"`,
                'tab': tab,
                'frozenRows': rows !== undefined ? parseInt( rows ) : null,
                'frozenCols': cols !== undefined ? parseInt( cols ) : null
            }

            return result
        } catch( err ) {
            return GetSheetCli.#error( { error: `Freeze failed: ${err.message}` } )
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


    static validationDelete( { tab, rows, cols } ) {
        const struct = { 'status': false, 'error': null }

        if( tab === undefined ) {
            struct['error'] = '--tab is required. Provide tab name'

            return struct
        }

        if( rows === undefined && cols === undefined ) {
            struct['error'] = 'One of --rows or --cols is required'

            return struct
        }

        if( rows !== undefined && cols !== undefined ) {
            struct['error'] = 'Only one of --rows or --cols can be specified'

            return struct
        }

        if( rows !== undefined && !/^\d+:\d+$/.test( rows ) ) {
            struct['error'] = '--rows must be in format "start:end", e.g. "2:5"'

            return struct
        }

        if( cols !== undefined && !/^[A-Za-z]+:[A-Za-z]+$/.test( cols ) ) {
            struct['error'] = '--cols must be in format "start:end", e.g. "B:C"'

            return struct
        }

        struct['status'] = true

        return struct
    }


    static validationFormat( { tab, range, bold, bg, color, fontsize, align } ) {
        const struct = { 'status': false, 'error': null }

        if( tab === undefined ) {
            struct['error'] = '--tab is required. Provide tab name'

            return struct
        }

        if( range === undefined ) {
            struct['error'] = '--range is required. e.g. A1:O1'

            return struct
        }

        const hasFormat = bold !== undefined || bg !== undefined || color !== undefined || fontsize !== undefined || align !== undefined
        if( !hasFormat ) {
            struct['error'] = 'At least one format option required: --bold, --bg, --color, --fontsize, --align'

            return struct
        }

        if( bg !== undefined && !/^#[0-9A-Fa-f]{6}$/.test( bg ) ) {
            struct['error'] = '--bg must be a hex color, e.g. "#4285f4"'

            return struct
        }

        if( color !== undefined && !/^#[0-9A-Fa-f]{6}$/.test( color ) ) {
            struct['error'] = '--color must be a hex color, e.g. "#333333"'

            return struct
        }

        const validAlignments = [ 'left', 'center', 'right' ]
        if( align !== undefined && !validAlignments.includes( align.toLowerCase() ) ) {
            struct['error'] = `--align must be one of: ${validAlignments.join( ', ' )}`

            return struct
        }

        struct['status'] = true

        return struct
    }


    static validationCondFormat( { tab, range, scale, gt, lt, eq, between, bg } ) {
        const struct = { 'status': false, 'error': null }

        if( tab === undefined ) {
            struct['error'] = '--tab is required. Provide tab name'

            return struct
        }

        if( range === undefined ) {
            struct['error'] = '--range is required. e.g. A1:N44'

            return struct
        }

        const hasCondition = gt !== undefined || lt !== undefined || eq !== undefined || between !== undefined

        if( scale === undefined && !hasCondition ) {
            struct['error'] = 'Either --scale or a condition (--gt, --lt, --eq, --between) is required'

            return struct
        }

        if( scale !== undefined && hasCondition ) {
            struct['error'] = 'Cannot combine --scale with conditions (--gt, --lt, --eq, --between)'

            return struct
        }

        if( scale !== undefined ) {
            const parts = scale.split( ':' )
            if( parts.length < 2 || parts.length > 3 ) {
                struct['error'] = '--scale must have 2 or 3 colors separated by ":", e.g. "red:green" or "red:yellow:green"'

                return struct
            }
        }

        if( hasCondition && bg === undefined ) {
            struct['error'] = '--bg is required when using conditions. Provide background color, e.g. "#4caf50"'

            return struct
        }

        if( bg !== undefined && !/^#[0-9A-Fa-f]{6}$/.test( bg ) ) {
            struct['error'] = '--bg must be a hex color, e.g. "#4caf50"'

            return struct
        }

        if( between !== undefined && !/^-?\d+(\.\d+)?:-?\d+(\.\d+)?$/.test( between ) ) {
            struct['error'] = '--between must be in format "min:max", e.g. "8:10"'

            return struct
        }

        struct['status'] = true

        return struct
    }


    static validationFreeze( { tab, rows, cols } ) {
        const struct = { 'status': false, 'error': null }

        if( tab === undefined ) {
            struct['error'] = '--tab is required. Provide tab name'

            return struct
        }

        if( rows === undefined && cols === undefined ) {
            struct['error'] = 'At least one of --rows or --cols is required'

            return struct
        }

        if( rows !== undefined && ( isNaN( parseInt( rows ) ) || parseInt( rows ) < 0 ) ) {
            struct['error'] = '--rows must be a non-negative number, e.g. "1"'

            return struct
        }

        if( cols !== undefined && ( isNaN( parseInt( cols ) ) || parseInt( cols ) < 0 ) ) {
            struct['error'] = '--cols must be a non-negative number, e.g. "1"'

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


    static #hexToRgb( { hex } ) {
        const cleaned = hex.replace( '#', '' )
        const bigint = parseInt( cleaned, 16 )
        const red = Math.round( ( ( bigint >> 16 ) & 255 ) / 255 * 1000 ) / 1000
        const green = Math.round( ( ( bigint >> 8 ) & 255 ) / 255 * 1000 ) / 1000
        const blue = Math.round( ( bigint & 255 ) / 255 * 1000 ) / 1000

        return { red, green, blue }
    }


    static #parseScale( { scale } ) {
        const namedColors = {
            'red': { 'red': 0.918, 'green': 0.263, 'blue': 0.208 },
            'green': { 'red': 0.204, 'green': 0.659, 'blue': 0.325 },
            'yellow': { 'red': 0.984, 'green': 0.737, 'blue': 0.016 },
            'white': { 'red': 1, 'green': 1, 'blue': 1 },
            'orange': { 'red': 1, 'green': 0.427, 'blue': 0.004 },
            'blue': { 'red': 0.259, 'green': 0.522, 'blue': 0.957 }
        }

        const colors = scale.split( ':' )
            .map( ( c ) => {
                const trimmed = c.trim().toLowerCase()
                if( namedColors[ trimmed ] ) {
                    const color = namedColors[ trimmed ]

                    return color
                }

                const { red, green, blue } = GetSheetCli.#hexToRgb( { hex: trimmed } )
                const color = { red, green, blue }

                return color
            } )

        return { colors }
    }


    static #error( { error } ) {
        const result = { 'status': false, error }

        return result
    }
}


export { GetSheetCli }
