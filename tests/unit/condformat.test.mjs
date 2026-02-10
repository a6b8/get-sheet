import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'


describe( 'GetSheetCli.validationCondFormat', () => {
    it( 'should return error when tab is missing', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: undefined, range: 'A1:B10', scale: 'red:green',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--tab' )
    } )


    it( 'should return error when range is missing', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: undefined, scale: 'red:green',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--range' )
    } )


    it( 'should return error when neither scale nor condition is provided', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--scale' )
    } )


    it( 'should return error when scale and condition are combined', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: 'red:green',
            gt: '5', lt: undefined, eq: undefined, between: undefined, bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'Cannot combine' )
    } )


    it( 'should return error when scale has less than 2 colors', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: 'red',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '2 or 3 colors' )
    } )


    it( 'should return error when scale has more than 3 colors', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: 'red:yellow:green:blue',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '2 or 3 colors' )
    } )


    it( 'should return error when condition used without bg', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: '5', lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--bg is required' )
    } )


    it( 'should return error when bg is invalid hex', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: '5', lt: undefined, eq: undefined, between: undefined, bg: '#xyz'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( 'hex color' )
    } )


    it( 'should return error when between has invalid format', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: undefined, lt: undefined, eq: undefined, between: 'abc', bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toContain( '--between' )
    } )


    it( 'should pass with valid 2-color scale', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: 'red:green',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with valid 3-color scale', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: 'red:yellow:green',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with valid gt condition and bg', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: '100', lt: undefined, eq: undefined, between: undefined, bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with valid between condition and bg', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: undefined,
            gt: undefined, lt: undefined, eq: undefined, between: '8:10', bg: '#c8e6c9'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'should pass with hex colors in scale', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1', range: 'A1:B10', scale: '#ff0000:#ffff00:#00ff00',
            gt: undefined, lt: undefined, eq: undefined, between: undefined, bg: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
