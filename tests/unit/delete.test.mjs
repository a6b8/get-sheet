import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB } from '../helpers/config.mjs'


describe( 'GetSheetCli.validationDelete', () => {
    it( 'returns error when tab is undefined', () => {
        const result = GetSheetCli.validationDelete( {
            tab: undefined,
            rows: '2:5',
            cols: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--tab is required. Provide tab name' )
    } )


    it( 'returns error when both rows and cols are undefined', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'One of --rows or --cols is required' )
    } )


    it( 'returns error when both rows and cols are provided', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: '2:5',
            cols: 'B:C'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Only one of --rows or --cols can be specified' )
    } )


    it( 'returns error when rows has invalid format', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: 'abc',
            cols: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--rows must be in format "start:end", e.g. "2:5"' )
    } )


    it( 'returns error when rows has letters instead of numbers', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: 'A:B',
            cols: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--rows must be in format "start:end", e.g. "2:5"' )
    } )


    it( 'returns error when cols has invalid format', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: '123'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--cols must be in format "start:end", e.g. "B:C"' )
    } )


    it( 'returns error when cols has numbers instead of letters', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: '1:3'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--cols must be in format "start:end", e.g. "B:C"' )
    } )


    it( 'returns status true for valid rows "2:5"', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: '2:5',
            cols: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for valid cols "B:C"', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: 'B:C'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for single row range "3:3"', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: '3:3',
            cols: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for single col range "A:A"', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: 'A:A'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for lowercase cols "b:d"', () => {
        const result = GetSheetCli.validationDelete( {
            tab: VALID_TAB,
            rows: undefined,
            cols: 'b:d'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
