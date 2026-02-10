import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_CREDENTIALS_PATH, VALID_SPREADSHEET_ID } from '../helpers/config.mjs'


describe( 'GetSheetCli.validationInit', () => {
    it( 'returns error when credentials is undefined', () => {
        const result = GetSheetCli.validationInit( {
            credentials: undefined,
            spreadsheet: VALID_SPREADSHEET_ID
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--credentials is required (first-time setup). Provide path to Google service account JSON' )
    } )


    it( 'returns error when spreadsheet is undefined', () => {
        const result = GetSheetCli.validationInit( {
            credentials: VALID_CREDENTIALS_PATH,
            spreadsheet: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--spreadsheet is required. Provide Google Spreadsheet ID' )
    } )


    it( 'returns status true when both are provided', () => {
        const result = GetSheetCli.validationInit( {
            credentials: VALID_CREDENTIALS_PATH,
            spreadsheet: VALID_SPREADSHEET_ID
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
