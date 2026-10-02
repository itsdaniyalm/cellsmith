import { describe, expect, it } from 'vitest';
import { isCellReference, tokenize } from '../src/core/tokenizer';

const sig = (src: string) => tokenize(src).filter((t) => t.kind !== 'whitespace').map((t) => [t.kind, t.text]);

describe('tokenizer', () => {
  it('tokenizes literals', () => {
    expect(sig('=1.5E-3+.5')).toEqual([
      ['operator', '='],
      ['number', '1.5E-3'],
      ['operator', '+'],
      ['number', '.5'],
    ]);
    expect(sig('"a ""b"" c"')).toEqual([['string', '"a ""b"" c"']]);
    expect(sig('TRUE FALSE true')).toEqual([['boolean', 'TRUE'], ['boolean', 'FALSE'], ['boolean', 'true']]);
    expect(sig('#N/A #DIV/0! #SPILL!')).toEqual([['error', '#N/A'], ['error', '#DIV/0!'], ['error', '#SPILL!']]);
  });

  it('marks unterminated strings', () => {
    const [t] = tokenize('"abc');
    expect(t?.kind).toBe('string');
    expect(t?.unterminated).toBe(true);
  });

  it('distinguishes cells, names and functions', () => {
    expect(sig('A1 $B$2 XFD1048576 XFE1 A0 myName LOG10(')).toEqual([
      ['cell', 'A1'],
      ['cell', '$B$2'],
      ['cell', 'XFD1048576'],
      ['name', 'XFE1'],
      ['name', 'A0'],
      ['name', 'myName'],
      ['function', 'LOG10'],
      ['lparen', '('],
    ]);
    expect(isCellReference('XFD1048576')).toBe(true);
    expect(isCellReference('XFE1')).toBe(false);
  });

  it('handles dotted function names and prefixes', () => {
    expect(sig('STDEV.S(')[0]).toEqual(['function', 'STDEV.S']);
    expect(sig('_xlfn.XLOOKUP(')[0]).toEqual(['function', '_xlfn.XLOOKUP']);
  });

  it('tokenizes whole column and row ranges', () => {
    expect(sig('A:A')).toEqual([['columnRange', 'A:A']]);
    expect(sig('$A:$C')).toEqual([['columnRange', '$A:$C']]);
    expect(sig('1:1')).toEqual([['rowRange', '1:1']]);
    expect(sig('SUM(2:5)').map((t) => t[0])).toEqual(['function', 'lparen', 'rowRange', 'rparen']);
    expect(sig('Sheet1!A:B')).toEqual([['columnRange', 'Sheet1!A:B']]);
  });

  it('tokenizes sheet-qualified references', () => {
    expect(sig('Sheet1!A1')).toEqual([['cell', 'Sheet1!A1']]);
    expect(sig("'My Sheet'!A1")).toEqual([['cell', "'My Sheet'!A1"]]);
    expect(sig("'It''s'!B2")).toEqual([['cell', "'It''s'!B2"]]);
    expect(sig('Sheet1:Sheet3!A1')).toEqual([['cell', 'Sheet1:Sheet3!A1']]);
    expect(sig('[Book1.xlsx]Sheet1!A1')).toEqual([['cell', '[Book1.xlsx]Sheet1!A1']]);
    expect(sig("'[Book 1.xlsx]Sheet1'!A1")).toEqual([['cell', "'[Book 1.xlsx]Sheet1'!A1"]]);
    expect(sig('Sheet1!MyName')).toEqual([['name', 'Sheet1!MyName']]);
    expect(tokenize('Sheet1!A1')[0]?.prefixLength).toBe(7);
  });

  it('keeps A1:B2 as cell, colon, cell', () => {
    expect(sig('A1:B2')).toEqual([['cell', 'A1'], ['colon', ':'], ['cell', 'B2']]);
  });

  it('tokenizes table references', () => {
    expect(sig('Table1[Col]')).toEqual([['tableRef', 'Table1[Col]']]);
    expect(sig('[@Col]')).toEqual([['tableRef', '[@Col]']]);
    expect(sig('Table1[[#This Row],[Col One]]')).toEqual([['tableRef', 'Table1[[#This Row],[Col One]]']]);
    expect(sig("Table1[Col '[x']]")).toEqual([['tableRef', "Table1[Col '[x']]"]]);
    expect(sig('Table1[#Headers]')).toEqual([['tableRef', 'Table1[#Headers]']]);
  });

  it('distinguishes spill operator from error literals', () => {
    expect(sig('A1#')).toEqual([['cell', 'A1'], ['operator', '#']]);
    expect(sig('FILTER(a,b)#').at(-1)).toEqual(['operator', '#']);
    expect(sig('#REF!')).toEqual([['error', '#REF!']]);
  });

  it('tokenizes multi-char operators', () => {
    expect(sig('<> <= >= < > = &').map((t) => t[1])).toEqual(['<>', '<=', '>=', '<', '>', '=', '&']);
  });

  it('keeps every character: tokens tile the input exactly', () => {
    const src = '=IF(A1>=10, "x""y", SUM(\n  Sheet2!B:B,\t{1,2;3,4}))#';
    expect(tokenize(src).map((t) => t.text).join('')).toBe(src);
  });

  it('flags unknown characters', () => {
    expect(sig('A1 ~ B1').map((t) => t[0])).toEqual(['cell', 'unknown', 'cell']);
  });
});
