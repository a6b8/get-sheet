import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'


describe( 'GetSheetCli.validationFreeze', () => {
    it( 'should return error when tab is missing', () => {
        const result = GetSheetCli.validationFreeze( { tab: undefined, rows: '1', cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--tab' )
    } )


    it( 'should return error when both rows and cols are missing', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: undefined, cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--rows' )
    } )


    it( 'should return error when rows is not a number', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: 'abc', cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'non-negative number' )
    } )


    it( 'should return error when rows is negative', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: '-1', cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'non-negative number' )
    } )


    it( 'should return error when cols is not a number', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: undefined, cols: 'xyz' } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'non-negative number' )
    } )


    it( 'should pass with valid rows only', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: '1', cols: undefined } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with valid cols only', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: undefined, cols: '2' } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with both rows and cols', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: '1', cols: '1' } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with 0 to unfreeze', () => {
        const result = GetSheetCli.validationFreeze( { tab: 'Sheet1', rows: '0', cols: '0' } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
