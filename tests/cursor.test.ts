import { describe, expect, it } from 'vitest';
import { completionContext, enclosingCall, nodePathAt } from '../src/core/cursor';
import { parse } from '../src/core/parser';

/** Use `|` in the source to mark the caret. */
function at(marked: string) {
  const offset = marked.indexOf('|');
  const text = marked.replace('|', '');
  return { text, offset };
}
const ctx = (marked: string) => {
  const { text, offset } = at(marked);
  return completionContext(text, offset);
};
const call = (marked: string) => {
  const { text, offset } = at(marked);
  const r = enclosingCall(parse(text), offset);
  return r && { name: r.fn.name, argIndex: r.argIndex, argc: r.argc };
};

describe('completionContext', () => {
  it('finds the word being typed', () => {
    expect(ctx('=SU|')).toMatchObject({ kind: 'word', word: 'SU', start: 1, end: 3 });
    expect(ctx('=IF(A1>0,VLO|)')).toMatchObject({ kind: 'word', word: 'VLO' });
    expect(ctx('=|')).toMatchObject({ kind: 'word', word: '' });
    expect(ctx('=SUM(|')).toMatchObject({ kind: 'word', word: '' });
  });
  it('replaces the whole word when the caret is mid-word', () => {
    expect(ctx('=SU|M(A1)')).toMatchObject({ kind: 'word', word: 'SU', start: 1, end: 4, nextChar: '(' });
  });
  it('supports dotted function names', () => {
    expect(ctx('=STDEV.|')).toMatchObject({ kind: 'word', word: 'STDEV.' });
  });
  it('offers nothing inside strings, after a sheet bang, or for non-formulas', () => {
    expect(ctx('="hel|lo"').kind).toBe('none');
    expect(ctx('="a""b|').kind).toBe('none');
    expect(ctx('=Sheet1!A|').kind).toBe('none');
    expect(ctx('=$A|').kind).toBe('none');
    expect(ctx('hello|').kind).toBe('none');
  });
  it('resumes after a closed string', () => {
    expect(ctx('="a"&SU|')).toMatchObject({ kind: 'word', word: 'SU' });
  });
  it('detects table specifiers', () => {
    expect(ctx('=SUM(Sales[|')).toMatchObject({ kind: 'table', table: 'Sales', start: 11 });
    expect(ctx('=SUM(Sales[Re|')).toMatchObject({ kind: 'table', table: 'Sales' });
    expect(ctx('=Sales[[#This Row],[Un|')).toMatchObject({ kind: 'table', table: 'Sales' });
    expect(ctx('=[@Un|')).toMatchObject({ kind: 'table', table: undefined });
    expect(ctx('=[@|')).toMatchObject({ kind: 'table', start: 3, end: 3 });
  });
  it('returns to normal completion after a table specifier closes', () => {
    expect(ctx('=Sales[Region]+SU|')).toMatchObject({ kind: 'word', word: 'SU' });
  });
});

describe('enclosingCall', () => {
  it('finds the innermost call and argument index', () => {
    expect(call('=SUM(A1,|')).toEqual({ name: 'SUM', argIndex: 1, argc: 2 });
    expect(call('=IF(A1>0,SUM(|')).toEqual({ name: 'SUM', argIndex: 0, argc: 1 });
    expect(call('=IF(SUM(A1,B1),|,3)')).toEqual({ name: 'IF', argIndex: 1, argc: 3 });
  });
  it('ignores commas inside nested calls, arrays and parentheses', () => {
    expect(call('=IF(SUM(1,2,3),{1,2,3},(1,2),|)')).toEqual({ name: 'IF', argIndex: 3, argc: 4 });
  });
  it('works while the call is still unclosed', () => {
    expect(call('=VLOOKUP(A1, B:C, |')).toEqual({ name: 'VLOOKUP', argIndex: 2, argc: 3 });
  });
  it('treats the caret right after the name or after the closing paren as outside', () => {
    expect(call('=SUM|(A1)')).toBeUndefined();
    expect(call('=SUM(A1)|')).toBeUndefined();
    expect(call('=A1+|')).toBeUndefined();
  });
  it('treats the caret just inside the parens as the first argument', () => {
    expect(call('=SUM(|)')).toEqual({ name: 'SUM', argIndex: 0, argc: 1 });
  });
  it('handles multi-line formulas', () => {
    expect(call('=IF(\n    A1>0,\n    |\n)')).toEqual({ name: 'IF', argIndex: 1, argc: 2 });
  });
});

describe('nodePathAt', () => {
  it('descends to the deepest node', () => {
    const { text, offset } = at('=SUM(A1,B|2)');
    const path = nodePathAt(parse(text).formula.body, offset);
    expect(path.map((n) => n.type)).toEqual(['Function', 'Reference']);
  });
});
