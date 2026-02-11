import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB } from '../helpers/config.mjs'


describe( 'GetSheetCli.validationFilter', () => {
    it( 'returns error when tab is undefined', () => {
        const result = GetSheetCli.validationFilter( {
            tab: undefined,
            range: 'A1:D100'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--tab is required. Provide tab name' )
    } )


    it( 'returns error when range is undefined', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--range is required. e.g. A1:D100' )
    } )


    it( 'returns error when range has invalid format "abc"', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'abc'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--range must be in format "A1:D100"' )
    } )


    it( 'returns error when range has no row numbers "A:D"', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'A:D'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--range must be in format "A1:D100"' )
    } )


    it( 'returns error when range is lowercase "a1:d100"', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'a1:d100'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--range must be in format "A1:D100"' )
    } )


    it( 'returns status true for valid tab and range', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'A1:D100'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for single column range "A1:A50"', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'A1:A50'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for wide range "A1:Z999"', () => {
        const result = GetSheetCli.validationFilter( {
            tab: VALID_TAB,
            range: 'A1:Z999'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
