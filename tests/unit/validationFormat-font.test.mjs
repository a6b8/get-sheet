import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'
import { VALID_TAB } from '../helpers/config.mjs'


const VALID_RANGE = 'A1:C10'


describe( 'GetSheetCli.validationFormat - font parameter', () => {
    it( 'returns status true for font only', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined,
            font: 'Arial'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for font combined with bold', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: true,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined,
            font: 'Courier New'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for font combined with bg', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: '#4285f4',
            color: undefined,
            fontsize: undefined,
            align: undefined,
            font: 'Verdana'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns status true for font combined with all options', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: true,
            bg: '#4285f4',
            color: '#ffffff',
            fontsize: '14',
            align: 'center',
            font: 'Georgia'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns error when no format options including font are provided', () => {
        const result = GetSheetCli.validationFormat( {
            tab: VALID_TAB,
            range: VALID_RANGE,
            bold: undefined,
            bg: undefined,
            color: undefined,
            fontsize: undefined,
            align: undefined,
            font: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'At least one format option required: --bold, --bg, --color, --fontsize, --align, --wrap, --font' )
    } )
} )
