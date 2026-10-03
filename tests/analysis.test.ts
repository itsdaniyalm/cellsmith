import { describe, expect, it } from 'vitest';
import { analyze, symbolsAt, type LintContext } from '../src/core/analysis';
import { activeParamIndex, allFunctions, checkArity, getFunction, suggestFunctions } from '../src/core/functions';

const codes = (src: string, ctx?: LintContext) => analyze(src, ctx).diagnostics.map((d) => d.code);

describe('function catalog', () => {
  it('has a broad catalog with sane arities', () => {
    expect(allFunctions().length).toBeGreaterThan(450);
    for (const f of allFunctions()) {
      expect(f.minArgs).toBeLessThanOrEqual(f.maxArgs);
      expect(f.name).toMatch(/^[A-Z][A-Z0-9.]*$/);
    }
  });

  it('parses signatures into required, optional and repeating parameters', () => {
    const iff = getFunction('if')!;
    expect(iff.params.map((p) => [p.name, p.optional])).toEqual([
      ['logical_test', false],
      ['value_if_true', true],
      ['value_if_false', true],
    ]);
    expect([iff.minArgs, iff.maxArgs]).toEqual([1, 3]);

    const sum = getFunction('SUM')!;
    expect([sum.minArgs, sum.maxArgs, sum.repeat]).toEqual([1, 255, 1]);

    const ifs = getFunction('IFS')!;
    expect([ifs.minArgs, ifs.repeat]).toEqual([2, 2]);
  });

  it('resolves prefixed and mixed-case names', () => {
    expect(getFunction('_xlfn.XLOOKUP')?.name).toBe('XLOOKUP');
    expect(getFunction('_xlfn._xlws.FILTER')?.name).toBe('FILTER');
    expect(getFunction('stdev.s')?.name).toBe('STDEV.S');
    expect(getFunction('NOPE')).toBeUndefined();
  });

  it('maps the active parameter for repeating groups', () => {
    const ifs = getFunction('IFS')!; // logical_test1, value_if_true1, [logical_test2, value_if_true2], ...
    expect([0, 1, 2, 3, 4, 5].map((i) => activeParamIndex(ifs, i, 6))).toEqual([0, 1, 2, 3, 2, 3]);
    const sum = getFunction('SUM')!;
    expect([0, 1, 2, 9].map((i) => activeParamIndex(sum, i, 10))).toEqual([0, 1, 1, 1]);
    const lambda = getFunction('LAMBDA')!;
    expect(activeParamIndex(lambda, 0, 3)).toBe(0);
    expect(activeParamIndex(lambda, 2, 3)).toBe(1);
  });

  it('suggests close function names', () => {
    expect(suggestFunctions('VLOKUP')).toContain('VLOOKUP');
    expect(suggestFunctions('SUMIF')).not.toContain('SUMIF');
    expect(suggestFunctions('CONCATENAT')).toContain('CONCATENATE');
  });

  it('checks arity rules', () => {
    expect(checkArity(getFunction('IF')!, 0)).toMatch(/expects 1 to 3/);
    expect(checkArity(getFunction('IF')!, 4)).toMatch(/expects 1 to 3/);
    expect(checkArity(getFunction('PI')!, 1)).toMatch(/takes no arguments/);
    expect(checkArity(getFunction('IFS')!, 3)).toMatch(/even/);
    expect(checkArity(getFunction('LET')!, 4)).toMatch(/odd/);
    expect(checkArity(getFunction('LET')!, 3)).toBeUndefined();
    expect(checkArity(getFunction('SUMIFS')!, 4)).toMatch(/odd/);
  });
});

describe('linter: functions', () => {
  it('accepts valid formulas', () => {
    for (const f of [
      '=SUM(A1:A10)',
      '=IF(A1>0,"a","b")',
      '=IF(A1,,2)',
      '=XLOOKUP(A1,B:B,C:C)',
      '=SUMIFS(A:A,B:B,"x",C:C,">1")',
      '=IFS(A1>1,"a",A1>0,"b")',
      '=LET(x,1,y,2,x+y)',
      '=LAMBDA(x,x+1)(5)',
      '=NOW()',
    ]) {
      expect(codes(f), f).toEqual([]);
    }
  });

  it('flags wrong argument counts at the right place', () => {
    const a = analyze('=IF()');
    expect(a.diagnostics[0]?.code).toBe('arity');

    const tooMany = analyze('=IF(1,2,3,4)');
    const d = tooMany.diagnostics.find((x) => x.code === 'arity')!;
    expect('=IF(1,2,3,4)'.slice(d.start, d.end)).toBe('4'); // the extra argument

    const tooFew = analyze('=VLOOKUP(A1,B:C)');
    const d2 = tooFew.diagnostics.find((x) => x.code === 'arity')!;
    expect('=VLOOKUP(A1,B:C)'.slice(d2.start, d2.end)).toBe(')');
  });

  it('flags unknown functions and suggests a fix', () => {
    const a = analyze('=VLOKUP(A1,B:C,2,0)');
    const d = a.diagnostics.find((x) => x.code === 'function.unknown')!;
    expect(d.message).toContain('VLOOKUP');
    expect(d.fix?.text).toBe('VLOOKUP');
    expect('=VLOKUP(A1,B:C,2,0)'.slice(d.start, d.end)).toBe('VLOKUP');
  });

  it('does not flag workbook-defined names and named LAMBDAs', () => {
    expect(codes('=MyFunc(1)', { definedNames: new Set(['myfunc']) })).toEqual([]);
    const named = new Map([['tax', { min: 1, max: 2 }]]);
    expect(codes('=TAX(1)', { namedLambdas: named, definedNames: new Set(['tax']) })).toEqual([]);
    expect(codes('=TAX(1,2,3)', { namedLambdas: named, definedNames: new Set(['tax']) })).toContain('arity.lambda');
  });

  it('does not run arity checks on syntactically broken calls', () => {
    expect(codes('=IF(1 2')).not.toContain('arity');
  });
});

describe('linter: LET and LAMBDA scoping', () => {
  it('resolves variables and parameters', () => {
    const ctx = { definedNames: new Set<string>() };
    expect(codes('=LET(x,1,y,x+1,y*2)', ctx)).toEqual([]);
    expect(codes('=LAMBDA(a,b,a+b)(1,2)', ctx)).toEqual([]);
  });

  it('reports undefined names only when workbook context is known', () => {
    expect(codes('=foo+1')).toEqual([]);
    expect(codes('=foo+1', { definedNames: new Set() })).toContain('name.undefined');
    expect(codes('=foo+1', { definedNames: new Set(['foo']) })).toEqual([]);
  });

  it('a LET value cannot see its own name or later names', () => {
    const ctx = { definedNames: new Set<string>() };
    expect(codes('=LET(x,x+1,x)', ctx)).toContain('name.undefined');
    expect(codes('=LET(x,y,y,2,x)', ctx)).toContain('name.undefined');
  });

  it('reports unused variables and parameters as faded hints', () => {
    const a = analyze('=LET(x,1,y,2,y)');
    const d = a.diagnostics.find((x) => x.code === 'name.unused')!;
    expect(d.unnecessary).toBe(true);
    expect(d.severity).toBe('hint');
    expect('=LET(x,1,y,2,y)'.slice(d.start, d.end)).toBe('x');
    expect(codes('=LAMBDA(a,b,a)(1,2)')).toContain('name.unused');
  });

  it('rejects names that look like cell references', () => {
    expect(codes('=LET(A1,5,A1)')).toContain('name.looks-like-cell');
  });

  it('warns on duplicate LET names and errors on duplicate parameters', () => {
    expect(codes('=LET(x,1,x,2,x)')).toContain('name.duplicate');
    expect(codes('=LAMBDA(a,a,a)(1,2)')).toContain('name.duplicate');
  });

  it('checks LET argument structure', () => {
    expect(codes('=LET(x,1)')).toContain('arity');
    expect(codes('=LET(x,1,y,x)')).toContain('arity');
  });

  it('checks arity of local lambdas and inline invocations', () => {
    expect(codes('=LET(f,LAMBDA(a,b,a+b),f(1))')).toContain('arity.lambda');
    expect(codes('=LET(f,LAMBDA(a,b,a+b),f(1,2))')).toEqual([]);
    expect(codes('=LAMBDA(a,b,a+b)(1)')).toContain('arity.lambda');
    expect(codes('=LAMBDA(a,[b],a)(1)')).toEqual(['name.unused']);
  });

  it('exposes visible symbols at a caret position', () => {
    const src = '=LET(x,1,y,2,x+y)';
    const a = analyze(src);
    expect(symbolsAt(a, src.indexOf('x+y') + 1).map((s) => s.name).sort()).toEqual(['x', 'y']);
    expect(symbolsAt(a, src.indexOf('y,2') + 2).map((s) => s.name)).toEqual(['x']);
    expect(symbolsAt(a, 1)).toEqual([]);
  });

  it('keeps symbols visible while a LET is still being typed', () => {
    const src = '=LET(x,1,';
    const a = analyze(src);
    expect(symbolsAt(a, src.length).map((s) => s.name)).toEqual(['x']);
  });
});

describe('linter: unreplaced snippet placeholders', () => {
  const ctx = { definedNames: new Set<string>() };

  it('flags a name that equals the parameter it sits in, as one error each (not as undefined names)', () => {
    const c = codes('=VLOOKUP(lookup_value,table_array,col_index_num)', ctx);
    expect(c.filter((x) => x === 'placeholder')).toHaveLength(3);
    expect(c).not.toContain('name.undefined');
  });
  it('flags placeholders from the LET and LAMBDA snippets', () => {
    expect(codes('=LET(name1,value1,calculation_or_name2)', ctx)).toContain('placeholder');
    expect(codes('=LAMBDA(parameter1,calculation)', ctx)).toContain('placeholder');
  });
  it('does not flag the LET name or LAMBDA parameter slots themselves', () => {
    expect(codes('=LET(name1,5,name1)', ctx)).not.toContain('placeholder');
    expect(codes('=LAMBDA(parameter1,parameter1+1)(2)', ctx)).not.toContain('placeholder');
  });
  it('only flags a name sitting in the matching slot', () => {
    expect(codes('=IF(value_if_true,1,2)', ctx)).not.toContain('placeholder'); // wrong slot: just an undefined name
    expect(codes('=IF(value_if_true,1,2)', ctx)).toContain('name.undefined');
  });
  it('respects names that are really defined', () => {
    expect(codes('=SUM(number1)', { definedNames: new Set(['number1']) })).toEqual([]);
    expect(codes('=LET(number1,5,SUM(number1))')).toEqual([]);
  });
  it('is an error so auto-apply holds off until the placeholder is filled', () => {
    const d = analyze('=SUM(number1)').diagnostics.find((x) => x.code === 'placeholder')!;
    expect(d.severity).toBe('error');
    expect(d.message).toContain("'number1'");
  });
});

describe('linter: misc rules', () => {
  it('warns about #REF!', () => {
    expect(codes('=A1+#REF!')).toContain('ref.deleted');
    expect(codes('=Sheet1!#REF!+1')).toContain('ref.deleted');
  });
  it('warns about division by a literal zero', () => {
    expect(codes('=A1/0')).toContain('math.divide-by-zero');
    expect(codes('=A1/B1')).toEqual([]);
  });
  it('explains the intersection operator', () => {
    expect(codes('=SUM(A1 B1)')).toContain('op.intersection');
  });
  it('flags cell-like names outside the grid', () => {
    expect(codes('=XFE1+1')).toContain('ref.out-of-grid');
  });
  it('validates array constants', () => {
    expect(codes('={1,2;3,4}')).toEqual([]);
    expect(codes('={1,2;3}')).toContain('array.ragged');
    expect(codes('={A1,2}')).toContain('array.not-constant');
    expect(codes('={-1,"a",TRUE,#N/A}')).toEqual([]);
  });
  it('enforces Excel limits', () => {
    expect(codes('=' + '1+'.repeat(5000) + '1')).toContain('limit.formula-length');
    const deep = '=' + 'ABS('.repeat(65) + '1' + ')'.repeat(65);
    expect(codes(deep)).toContain('limit.nesting');
    expect(codes('=' + 'ABS('.repeat(64) + '1' + ')'.repeat(64))).not.toContain('limit.nesting');
    expect(codes('=SUM(' + Array(256).fill('1').join(',') + ')')).toContain('arity');
  });
  it('does not lint plain text that is not a formula', () => {
    expect(analyze('hello world (').diagnostics).toEqual([]);
  });
  it('includes syntax errors with positions', () => {
    const a = analyze('=SUM(A1,B1');
    expect(a.diagnostics.map((d) => d.code)).toContain('syntax.unclosed-paren');
  });
});

describe('linter: structured references', () => {
  it('rejects empty brackets with no table name, but allows Table[]', () => {
    expect(codes('=ROWS([])')).toContain('ref.empty-structured');
    expect(codes('=ROWS([ ])')).toContain('ref.empty-structured');
    expect(codes('=ROWS(Sales[])')).not.toContain('ref.empty-structured');
    expect(codes('=[@Units]*2')).not.toContain('ref.empty-structured');
    expect(codes('=LAMBDA(x, [y], x)(1)')).not.toContain('ref.empty-structured');
    const d = analyze('=ROWS([])').diagnostics.find((x) => x.code === 'ref.empty-structured')!;
    expect(d.severity).toBe('error');
  });
});
