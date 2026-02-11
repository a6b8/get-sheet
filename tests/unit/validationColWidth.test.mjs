import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB } from '../helpers/config.mjs'


describe( 'GetSheetCli.validationColWidth', () => {
    it( 'returns error when tab is undefined', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: undefined,
            cols: 'A',
            width: 150
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--tab is required. Provide tab name' )
    } )


    it( 'returns error when cols is undefined', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: undefined,
            width: 150
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--cols is required. e.g. "A" or "A:C"' )
    } )


    it( 'returns error when cols has invalid format "123"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: '123',
            width: 150
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--cols must be a single column (e.g. "A") or range (e.g. "A:C")' )
    } )


    it( 'returns error when cols has invalid format "A1:C3"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A1:C3',
            width: 150
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--cols must be a single column (e.g. "A") or range (e.g. "A:C")' )
    } )


    it( 'returns error when width is undefined', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A',
            width: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--width is required. Provide pixel value, e.g. 150' )
    } )


    it( 'returns error when width is 0', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A',
            width: 0
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--width must be a positive number (pixels)' )
    } )


    it( 'returns error when width is negative', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A',
            width: -1
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--width must be a positive number (pixels)' )
    } )


    it( 'returns error when width is not a number "abc"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A',
            width: 'abc'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--width must be a positive number (pixels)' )
    } )


    it( 'returns status true for single column "A"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A',
            width: 150
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for column range "A:C"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'A:C',
            width: 200
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for lowercase column "a"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'a',
            width: 100
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for lowercase column range "a:c"', () => {
        const result = GetSheetCli.validationColWidth( {
            tab: VALID_TAB,
            cols: 'a:c',
            width: 100
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
