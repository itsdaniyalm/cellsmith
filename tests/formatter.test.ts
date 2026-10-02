import { describe, expect, it } from 'vitest';
import { formatFormula, minifyFormula } from '../src/core/formatter';
import { parse } from '../src/core/parser';

const fmt = (s: string, o = {}) => formatFormula(s, o).text;

/** A spread of realistic formulas, including awkward syntax. */
const CORPUS = [
  '=A1+B1',
  '=SUM(A1:A10)',
  '=IF(A1>0,"positive",IF(A1<0,"negative","zero"))',
  '=IFERROR(VLOOKUP(A2,Sheet2!$A$1:$D$500,3,FALSE),"not found")',
  '=XLOOKUP(A2,Table1[ID],Table1[Name],"n/a",0,1)',
  '=SUMIFS(Sales[Amount],Sales[Region],F2,Sales[Date],">="&G2,Sales[Date],"<="&H2)',
  '=LET(x,A1:A10,y,SUM(x),z,y/COUNT(x),IF(z>10,"high","low"))',
  '=LET(total,SUM(Table1[Sales]),avg,AVERAGE(Table1[Sales]),total/avg)',
  '=LAMBDA(x,y,x^2+y^2)(3,4)',
  '=LET(sq,LAMBDA(n,n*n),sq(4)+sq(5))',
  '=MAP(A1:A10,LAMBDA(v,IF(ISNUMBER(v),v*2,"")))',
  '=REDUCE(0,A1:A10,LAMBDA(acc,v,acc+v))',
  '=FILTER(A2:C100,(B2:B100="x")*(C2:C100>5),"none")',
  '=SORT(UNIQUE(FILTER(A2:A100,A2:A100<>"")),1,-1)',
  '=TEXTJOIN(", ",TRUE,IF(B2:B20="yes",A2:A20,""))',
  '=INDEX(A:A,MATCH(MAX(B:B),B:B,0))',
  '=SUM(A1:INDEX(A:A,COUNTA(A:A)))',
  '=-2^2',
  '=2^-3',
  '=50%*A1',
  '=A1&" "&B1&" "&C1',
  '=SUM(A1#)',
  '=@A1:A5',
  '=SUM((A1:A3,C1:C3))',
  '=SUM(A1:B2 B2:C3)',
  '=IF(A1,,2)',
  '=INDEX({1,2,3;4,5,6},2,3)',
  '={"a","b";"c","d"}',
  "='My Sheet'!A1+'Other Sheet'!B2",
  '=Table1[[#This Row],[Qty]]*Table1[[#This Row],[Price]]',
  '=[@Qty]*[@Price]',
  '=NOW()-TODAY()',
  '=SUMPRODUCT((A2:A100="x")*(B2:B100)*(C2:C100>=DATE(2024,1,1)))',
  '=IF(AND(A1>0,OR(B1="x",B1="y"),NOT(ISBLANK(C1))),"ok","bad")',
  '=A1=B1',
  '=A1<>B1',
  '=((A1+B1)*(C1-D1))/E1',
  '=#REF!+1',
  '=_xlfn.XLOOKUP(A1,B:B,C:C)',
  '=sum(a1:a3)',
];

describe('formatter', () => {
  it('formats short formulas on one line with normalized spacing', () => {
    expect(fmt('=SUM(A1:A10)')).toBe('=SUM(A1:A10)');
    expect(fmt('=  IF( A1>0 ,"a" ,"b" )')).toBe('=IF(A1 > 0, "a", "b")');
    expect(fmt('=A1+B1*2')).toBe('=A1 + B1 * 2');
    expect(fmt('=-A1')).toBe('=-A1');
    expect(fmt('=50%')).toBe('=50%');
  });

  it('upper-cases known functions and cell refs but not names or strings', () => {
    expect(fmt('=sum(a1:a3)')).toBe('=SUM(A1:A3)');
    expect(fmt('=myFunc(myName)')).toBe('=myFunc(myName)');
    expect(fmt('=if(true,"abc","def")')).toBe('=IF(TRUE, "abc", "def")');
  });

  it('breaks long calls one argument per line', () => {
    const out = fmt('=IFERROR(VLOOKUP(A2,Sheet2!$A$1:$D$500,3,FALSE),"a long fallback message that does not fit")', { width: 60 });
    expect(out).toBe(
      ['=IFERROR(', '    VLOOKUP(A2, Sheet2!$A$1:$D$500, 3, FALSE),', '    "a long fallback message that does not fit"', ')'].join('\n'),
    );
  });

  it('expands LET with several bindings and keeps name/value pairs together', () => {
    expect(fmt('=LET(x,A1:A10,y,SUM(x),y/COUNT(x))')).toBe(
      ['=LET(', '    x, A1:A10,', '    y, SUM(x),', '    y / COUNT(x)', ')'].join('\n'),
    );
  });

  it('keeps a single-binding LET flat when it fits', () => {
    expect(fmt('=LET(x,5,x*2)')).toBe('=LET(x, 5, x * 2)');
  });

  it('indents nested structures inside LET values', () => {
    const out = fmt('=LET(rows,FILTER(A2:C100,(B2:B100="x")*(C2:C100>5),"none"),n,ROWS(rows),n)', { width: 50 });
    expect(out).toBe(
      [
        '=LET(',
        '    rows, FILTER(',
        '        A2:C100,',
        '        (B2:B100 = "x") * (C2:C100 > 5),',
        '        "none"',
        '    ),',
        '    n, ROWS(rows),',
        '    n',
        ')',
      ].join('\n'),
    );
  });

  it('keeps LAMBDA parameters together and puts the body below when too long', () => {
    const out = fmt('=LAMBDA(a,b,c,IF(a>b,SUM(a,b,c)*1000000,AVERAGE(a,b,c)/1000000))', { width: 40 });
    expect(out.split('\n')[1]).toBe('    a, b, c,');
  });

  it('breaks long operator chains before the operator', () => {
    const out = fmt('=A1&" some fairly long text "&B1&" and some more text "&C1', { width: 30 });
    expect(out).toBe(['=A1', '    & " some fairly long text "', '    & B1', '    & " and some more text "', '    & C1'].join('\n'));
  });

  it('respects the indent option', () => {
    expect(fmt('=IF(A1>0,"aaaaaaaaaaaa","bbbbbbbbbbbbbb")', { width: 20, indent: 2 })).toBe('=IF(\n  A1 > 0,\n  "aaaaaaaaaaaa",\n  "bbbbbbbbbbbbbb"\n)');
  });

  it('refuses to touch formulas with syntax errors or non-formulas', () => {
    expect(formatFormula('=SUM(A1,').ok).toBe(false);
    expect(fmt('=SUM(A1,')).toBe('=SUM(A1,');
    expect(formatFormula('hello').ok).toBe(false);
  });

  it('never alters string literals, even with newlines inside', () => {
    expect(fmt('="line1\nline2"&A1')).toBe('="line1\nline2" & A1');
  });

  it('preserves the space intersection operator and union', () => {
    expect(fmt('=SUM(A1:B2   B2:C3)')).toBe('=SUM(A1:B2 B2:C3)');
    expect(fmt('=SUM((A1:A3 , C1:C3))')).toBe('=SUM((A1:A3,C1:C3))');
  });

  describe.each(CORPUS)('corpus: %s', (src) => {
    it('keeps the formula meaning (minified form is identical)', () => {
      const out = formatFormula(src);
      expect(out.ok).toBe(true);
      expect(minifyFormula(out.text).text).toBe(minifyFormula(src).text);
    });
    it('is idempotent', () => {
      const once = fmt(src);
      expect(fmt(once)).toBe(once);
    });
    it('produces output that parses without errors, at several widths', () => {
      for (const width of [20, 40, 80, 200]) {
        const out = fmt(src, { width });
        expect(parse(out).diagnostics).toEqual([]);
        expect(minifyFormula(out).text).toBe(minifyFormula(src).text);
      }
    });
    it('minify is stable', () => {
      const m = minifyFormula(src).text;
      expect(minifyFormula(m).text).toBe(m);
      expect(m).not.toMatch(/\n/);
    });
  });
});
