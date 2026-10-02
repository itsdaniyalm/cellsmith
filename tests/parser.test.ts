import { describe, expect, it } from 'vitest';
import type { Node } from '../src/core/ast';
import { parse } from '../src/core/parser';

/** Compact S-expression of the AST for readable assertions. */
function sx(n: Node): string {
  switch (n.type) {
    case 'Number':
    case 'String':
    case 'Boolean':
    case 'ErrorLiteral':
      return n.raw;
    case 'Reference':
      return n.text;
    case 'Missing':
      return '_';
    case 'Invalid':
      return `!${n.raw}`;
    case 'Unary':
      return `(${n.op}u ${sx(n.operand)})`;
    case 'Postfix':
      return `(${sx(n.operand)} ${n.op}p)`;
    case 'Paren':
      return `(paren ${sx(n.expr)})`;
    case 'Array':
      return `{${n.rows.map((r) => r.map(sx).join(',')).join(';')}}`;
    case 'Function':
      return `${n.name}[${n.args.map(sx).join(' | ')}]`;
    case 'Call':
      return `call(${sx(n.callee)})[${n.args.map(sx).join(' | ')}]`;
    case 'Binary':
      return `(${sx(n.left)} ${n.op === ' ' ? '∩' : n.op} ${sx(n.right)})`;
  }
}
const ast = (src: string) => sx(parse(src).formula.body);
const errs = (src: string) => parse(src).diagnostics.map((d) => d.code);

describe('parser: precedence and associativity', () => {
  it('multiplication binds tighter than addition', () => {
    expect(ast('=1+2*3')).toBe('(1 + (2 * 3))');
    expect(ast('=1*2+3')).toBe('((1 * 2) + 3)');
  });
  it('is left-associative, including ^', () => {
    expect(ast('=1-2-3')).toBe('((1 - 2) - 3)');
    expect(ast('=2^3^2')).toBe('((2 ^ 3) ^ 2)');
  });
  it('unary minus binds tighter than ^ (Excel quirk: =-2^2 is 4)', () => {
    expect(ast('=-2^2')).toBe('((-u 2) ^ 2)');
  });
  it('percent is postfix', () => {
    expect(ast('=50%*2')).toBe('((50 %p) * 2)');
    expect(ast('=1+2%')).toBe('(1 + (2 %p))');
  });
  it('& is below + and above comparison', () => {
    expect(ast('="a"&1+2')).toBe('("a" & (1 + 2))');
    expect(ast('=A1&"x"=B1')).toBe('((A1 & "x") = B1)');
  });
  it('range binds tighter than everything', () => {
    expect(ast('=SUM(A1:B2)+1')).toBe('(SUM[(A1 : B2)] + 1)');
    expect(ast('=-A1:B2')).toBe('(-u (A1 : B2))');
  });
  it('supports ranges built from functions', () => {
    expect(ast('=SUM(A1:INDEX(A:A,5))')).toBe('SUM[(A1 : INDEX[A:A | 5])]');
  });
});

describe('parser: references and operators', () => {
  it('parses intersection (space) between references', () => {
    expect(ast('=SUM(A1:B2 B2:C3)')).toBe('SUM[((A1 : B2) ∩ (B2 : C3))]');
    expect(ast('=Sales Q1')).toBe('(Sales ∩ Q1)');
  });
  it('does not treat a newline as intersection', () => {
    expect(errs('=A1\nB1')).toContain('syntax.unexpected-token');
  });
  it('parses union inside parentheses only', () => {
    expect(ast('=SUM((A1:A2,C1:C2))')).toBe('SUM[(paren ((A1 : A2) , (C1 : C2)))]');
    expect(ast('=SUM(A1,B1)')).toBe('SUM[A1 | B1]');
  });
  it('parses spill, implicit intersection and table refs', () => {
    expect(ast('=SUM(A1#)')).toBe('SUM[(A1 #p)]');
    expect(ast('=@A1:A5')).toBe('(@u (A1 : A5))');
    expect(ast('=Table1[[#This Row],[Sales]]*2')).toBe('(Table1[[#This Row],[Sales]] * 2)');
    expect(ast('=[@Qty]*[@Price]')).toBe('([@Qty] * [@Price])');
  });
  it('parses sheet-qualified references', () => {
    expect(ast("='My Sheet'!A1:B2")).toBe("('My Sheet'!A1 : B2)");
    expect(ast('=Sheet1!A:A')).toBe('Sheet1!A:A');
  });
  it('treats #REF! inside a reference as an error reference', () => {
    const f = parse('=Sheet1!#REF!').formula.body;
    expect(f.type === 'Reference' && f.kind).toBe('error');
  });
});

describe('parser: functions', () => {
  it('parses empty and missing arguments', () => {
    expect(ast('=NOW()')).toBe('NOW[]');
    expect(ast('=IF(A1,,2)')).toBe('IF[A1 | _ | 2]');
    expect(ast('=SUM(1,)')).toBe('SUM[1 | _]');
    expect(errs('=IF(A1,,2)')).toEqual([]);
  });
  it('parses nested calls and array constants', () => {
    expect(ast('=INDEX({1,2;3,4},2,1)')).toBe('INDEX[{1,2;3,4} | 2 | 1]');
    expect(ast('=IF(A1>0,"y",IF(A1<0,"n","z"))')).toBe('IF[(A1 > 0) | "y" | IF[(A1 < 0) | "n" | "z"]]');
  });
  it('parses immediate LAMBDA invocation', () => {
    expect(ast('=LAMBDA(x,x+1)(5)')).toBe('call(LAMBDA[x | (x + 1)])[5]');
    expect(errs('=LAMBDA(x,x+1)(5)')).toEqual([]);
  });
  it('allows whitespace and newlines between tokens', () => {
    expect(ast('=IF(\n  A1 > 0,\n  "y",\n  "n"\n)')).toBe('IF[(A1 > 0) | "y" | "n"]');
  });
  it('parses a body without a leading equals sign', () => {
    const r = parse('1+2');
    expect(r.formula.hasEquals).toBe(false);
    expect(sx(r.formula.body)).toBe('(1 + 2)');
  });
});

describe('parser: errors', () => {
  it('reports unterminated strings', () => {
    expect(errs('="abc')).toContain('syntax.unterminated-string');
  });
  it('reports missing closing parens at the opening paren', () => {
    const r = parse('=SUM(A1,B1');
    expect(r.diagnostics.map((d) => d.code)).toContain('syntax.unclosed-paren');
    expect(r.diagnostics[0]!.start).toBe(4);
  });
  it('reports unmatched closing parens', () => {
    expect(errs('=SUM(A1))')).toContain('syntax.unmatched-paren');
  });
  it('reports dangling operators', () => {
    expect(errs('=1+')).toContain('syntax.expected-expression');
    expect(errs('=*2')).toContain('syntax.expected-expression');
    expect(errs('=')).toContain('syntax.expected-expression');
  });
  it('reports junk between arguments and recovers', () => {
    const r = parse('=SUM(1 2, 3)');
    expect(r.diagnostics.length).toBeGreaterThan(0);
    // The tree is still produced so completion/signature help keep working.
    expect(r.formula.body.type).toBe('Function');
  });
  it('reports bad characters', () => {
    expect(errs('=A1~B1')).toContain('syntax.bad-token');
  });
  it('never throws or loops on arbitrary garbage', () => {
    const junk = ['=)', '=(', '=((((', '=}{', '={', '={1,2', '=,,,', '=:::', "='", '="', '=[', '=[[[', '=#', '=@@@', '=SUM(', '=SUM(,', '=%%', '=A1:'];
    for (const j of junk) expect(() => parse(j)).not.toThrow();
  });
});
