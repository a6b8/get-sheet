import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB, VALID_DATA_2D } from '../helpers/config.mjs'


describe( 'GetSheetCli.validationWrite', () => {
    it( 'returns error when tab is undefined', () => {
        const result = GetSheetCli.validationWrite( {
            tab: undefined,
            data: VALID_DATA_2D
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--tab is required. Provide tab name' )
    } )


    it( 'returns error when data is undefined', () => {
        const result = GetSheetCli.validationWrite( {
            tab: VALID_TAB,
            data: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--data is required. Provide 2D array as JSON' )
    } )


    it( 'returns error when data is not an array', () => {
        const result = GetSheetCli.validationWrite( {
            tab: VALID_TAB,
            data: 'not an array'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--data must be a 2D array. Example: [["a","b"],["c","d"]]' )
    } )


    it( 'returns error when data is empty array', () => {
        const result = GetSheetCli.validationWrite( {
            tab: VALID_TAB,
            data: []
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--data must not be empty' )
    } )


    it( 'returns error when data contains non-array rows', () => {
        const result = GetSheetCli.validationWrite( {
            tab: VALID_TAB,
            data: [ 'not a row', 'also not a row' ]
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--data must be a 2D array. Each row must be an array' )
    } )


    it( 'returns status true for valid tab and data', () => {
        const result = GetSheetCli.validationWrite( {
            tab: VALID_TAB,
            data: VALID_DATA_2D
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
