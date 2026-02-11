import { describe, it, expect } from '@jest/globals'

import { GetSheetCli } from '../../src/GetSheetCli.mjs'


describe( 'GetSheetCli.validationCondFormat - formula parameter', () => {
    it( 'returns status true for formula with bg', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns error when formula is combined with scale', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: 'red:green',
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Cannot combine --scale with --formula' )
    } )


    it( 'returns error when formula is combined with gt', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: '5',
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Cannot combine --formula with conditions (--gt, --lt, --eq, --between)' )
    } )


    it( 'returns error when formula is combined with lt', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: '10',
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Cannot combine --formula with conditions (--gt, --lt, --eq, --between)' )
    } )


    it( 'returns error when formula is combined with eq', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: '42',
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Cannot combine --formula with conditions (--gt, --lt, --eq, --between)' )
    } )


    it( 'returns error when formula is combined with between', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: '1:10',
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Cannot combine --formula with conditions (--gt, --lt, --eq, --between)' )
    } )


    it( 'returns error when formula is used without bg', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( '--bg is required when using conditions or --formula. Provide background color, e.g. "#4caf50"' )
    } )


    it( 'returns status true for formula with bg and bold', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: '=A1>100',
            bg: '#4caf50'
        } )

        expect( result['status'] ).toBe( true )
        expect( result['error'] ).toBeNull()
    } )


    it( 'returns error when neither scale nor formula nor condition is provided', () => {
        const result = GetSheetCli.validationCondFormat( {
            tab: 'Sheet1',
            range: 'A1:B10',
            scale: undefined,
            gt: undefined,
            lt: undefined,
            eq: undefined,
            between: undefined,
            formula: undefined,
            bg: undefined
        } )

        expect( result['status'] ).toBe( false )
        expect( result['error'] ).toBe( 'Either --scale, --formula, or a condition (--gt, --lt, --eq, --between) is required' )
    } )
} )
