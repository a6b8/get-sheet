import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB } from '../helpers/config.mjs'


const VALID_RANGE = 'A1:C10'


describe( 'GetSheetCli.validationFormat', () => {
    it( 'returns error when tab is undefined', () => {
        const result = GetSheetCli.validationFormat( {
            tab: undefined,
            range: VALID_RANGE,
            bold: true,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--tab is required. Provide tab name' )
    } )


    it( 'returns error when range is undefined', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: undefined,
            bold: true,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--range is required. e.g. A1:O1' )
    } )


    it( 'returns error when no format options are provided', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'At least one format option required: --bold, --bg, --color, --fontsize, --align, --font' )
    } )


    it( 'returns error when bg is invalid hex', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: '#xyz',
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--bg must be a hex color, e.g. "#4285f4"' )
    } )


    it( 'returns error when bg hex is too short', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: '#fff',
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--bg must be a hex color, e.g. "#4285f4"' )
    } )


    it( 'returns error when color is invalid hex', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: 'red',
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--color must be a hex color, e.g. "#333333"' )
    } )


    it( 'returns error when color hex has no hash prefix', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: '333333',
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--color must be a hex color, e.g. "#333333"' )
    } )


    it( 'returns error when align is invalid value', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: 'middle'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--align must be one of: left, center, right' )
    } )


    it( 'returns error when align is "top"', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: 'top'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--align must be one of: left, center, right' )
    } )


    it( 'returns status true for bold only', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: true,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for valid bg "#4285f4"', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: '#4285f4',
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for valid color "#333333"', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: '#333333',
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for valid fontsize', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: '14',
            align: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for valid align "center"', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: 'center'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for align "LEFT" (case insensitive)', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: 'LEFT'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for multiple format options', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: true,
            bg: '#4285f4',
            color: '#ffffff',
            fontsize: '12',
            align: 'right'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for bg with uppercase hex "#AABBCC"', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: '#AABBCC',
            color: undefined,
            fontsize: undefined,
            align: undefined
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )
} )
