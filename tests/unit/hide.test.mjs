import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'


describe( 'GetSheetCli.validationHide', () => {
    it( 'should return error when tab is missing', () => {
        const result = GetSheetCli.validationHide( { tab: undefined, rows: '2:5', cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--tab' )
    } )


    it( 'should return error when both rows and cols are missing', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: undefined, cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--rows' )
    } )


    it( 'should return error when both rows and cols are provided', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: '2:5', cols: 'B:C' } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'Only one' )
    } )


    it( 'should return error when rows format is invalid', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: 'abc', cols: undefined } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'format' )
    } )


    it( 'should return error when cols format is invalid', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: undefined, cols: '123' } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'format' )
    } )


    it( 'should pass with valid rows', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: '2:5', cols: undefined } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with valid cols', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: undefined, cols: 'B:C' } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with single row range', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: '3:3', cols: undefined } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with single column', () => {
        const result = GetSheetCli.validationHide( { tab: 'Sheet1', rows: undefined, cols: 'A:A' } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
