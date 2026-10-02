import type { Diagnostic, Formula, Node, Reference, Span } from './ast';
import { tokenize, type Token } from './tokenizer';

export interface ParseResult {
  formula: Formula;
  tokens: Token[];
  diagnostics: Diagnostic[];
}

/** Binding power; higher binds tighter. Mirrors Excel's documented operator precedence. */
const PREC = {
  comparison: 1,
  concat: 2,
  additive: 3,
  multiplicative: 4,
  power: 5,
  percent: 6,
  unary: 7,
  union: 8,
  intersect: 9,
  range: 10,
} as const;

const BINARY_PREC: Record<string, number> = {
  '=': PREC.comparison,
  '<>': PREC.comparison,
  '<': PREC.comparison,
  '>': PREC.comparison,
  '<=': PREC.comparison,
  '>=': PREC.comparison,
  '&': PREC.concat,
  '+': PREC.additive,
  '-': PREC.additive,
  '*': PREC.multiplicative,
  '/': PREC.multiplicative,
  '^': PREC.power,
};

interface SigToken {
  tok: Token;
  /** Whitespace immediately before this token. */
  ws: string;
}

const INTERSECT_OPERAND = new Set(['cell', 'columnRange', 'rowRange', 'name', 'tableRef', 'function', 'lparen']);

function isReferenceLike(n: Node): boolean {
  switch (n.type) {
    case 'Reference':
      return n.kind !== 'error';
    case 'Function':
    case 'Paren':
      return true;
    case 'Binary':
      return n.op === ':' || n.op === ' ' || n.op === ',';
    case 'Postfix':
      return n.op === '#';
    default:
      return false;
  }
}

export function parse(src: string): ParseResult {
  const tokens = tokenize(src);
  const sig: SigToken[] = [];
  let pendingWs = '';
  for (const tok of tokens) {
    if (tok.kind === 'whitespace') pendingWs += tok.text;
    else {
      sig.push({ tok, ws: pendingWs });
      pendingWs = '';
    }
  }

  const diagnostics: Diagnostic[] = [];
  let p = 0;
  let allowUnion = false;

  const peek = (): Token | undefined => sig[p]?.tok;
  const peekWs = (): string => sig[p]?.ws ?? '';
  const next = (): Token => sig[p++]!.tok;
  const eofPos = (): number => src.length;
  const here = (): number => peek()?.start ?? eofPos();
  const error = (start: number, end: number, code: string, message: string) =>
    diagnostics.push({ severity: 'error', code, message, start, end });

  const isOp = (t: Token | undefined, text: string) => !!t && t.kind === 'operator' && t.text === text;

  let hasEquals = false;
  if (peek() && isOp(peek(), '=') && peek()!.start === 0) {
    hasEquals = true;
    next();
  }

  function missing(): Node {
    const at = here();
    return { type: 'Missing', start: at, end: at };
  }

  function describe(t: Token): string {
    return t.kind === 'rparen' ? "')'" : `'${t.text}'`;
  }

  /** Skip tokens until one of the stop kinds at the current nesting depth (does not consume the stop). */
  function skipTo(stops: ReadonlySet<string>): Node | undefined {
    const start = here();
    let depth = 0;
    let end = start;
    while (peek()) {
      const t = peek()!;
      if (depth === 0 && stops.has(t.kind)) break;
      if (t.kind === 'lparen' || t.kind === 'lbrace') depth++;
      if ((t.kind === 'rparen' || t.kind === 'rbrace') && depth > 0) depth--;
      end = t.end;
      next();
    }
    return end > start ? { type: 'Invalid', raw: src.slice(start, end), start, end } : undefined;
  }

  function parseExpression(minPrec: number): Node {
    let left = parseUnary();
    for (;;) {
      const t = peek();
      if (!t) break;

      // Postfix percent
      if (isOp(t, '%')) {
        if (PREC.percent < minPrec) break;
        next();
        left = { type: 'Postfix', op: '%', operand: left, start: left.start, end: t.end };
        continue;
      }

      let op: string | undefined;
      let prec = 0;
      let consumeOp = true;
      if (t.kind === 'operator' && BINARY_PREC[t.text] !== undefined) {
        op = t.text;
        prec = BINARY_PREC[t.text]!;
      } else if (t.kind === 'colon') {
        op = ':';
        prec = PREC.range;
      } else if (t.kind === 'comma' && allowUnion) {
        op = ',';
        prec = PREC.union;
      } else if (
        INTERSECT_OPERAND.has(t.kind) &&
        peekWs() !== '' &&
        !peekWs().includes('\n') &&
        isReferenceLike(left)
      ) {
        op = ' ';
        prec = PREC.intersect;
        consumeOp = false;
      }
      if (op === undefined || prec < minPrec) break;

      let opSpan: Span;
      if (consumeOp) {
        next();
        opSpan = { start: t.start, end: t.end };
      } else {
        const wsLen = peekWs().length;
        opSpan = { start: t.start - wsLen, end: t.start };
      }
      const right = parseExpression(prec + 1);
      left = { type: 'Binary', op, opSpan, left, right, start: left.start, end: Math.max(right.end, opSpan.end) };
    }
    return left;
  }

  function parseUnary(): Node {
    const t = peek();
    if (t && t.kind === 'operator' && (t.text === '-' || t.text === '+')) {
      next();
      const operand = parseExpression(PREC.unary);
      return { type: 'Unary', op: t.text as '-' | '+', operand, start: t.start, end: operand.end };
    }
    if (t && isOp(t, '@')) {
      next();
      const operand = parseExpression(PREC.intersect);
      return { type: 'Unary', op: '@', operand, start: t.start, end: operand.end };
    }
    return parsePostfixSpill(parsePrimary());
  }

  /** `A1#`, `FILTER(...)#` and `(expr)#` spill references, plus `LAMBDA(...)(args)` invocation. */
  function parsePostfixSpill(node: Node): Node {
    for (;;) {
      const t = peek();
      if (t && isOp(t, '#') && peekWs() === '') {
        next();
        node = { type: 'Postfix', op: '#', operand: node, start: node.start, end: t.end };
      } else if (t && t.kind === 'lparen' && peekWs() === '' && (node.type === 'Function' || node.type === 'Call' || node.type === 'Paren')) {
        next();
        const { args, closed, end } = parseArgs(t, 'this call');
        node = { type: 'Call', callee: node, args, closed, start: node.start, end };
      } else return node;
    }
  }

  function parsePrimary(): Node {
    const t = peek();
    if (!t) {
      error(eofPos(), eofPos(), 'syntax.expected-expression', 'Unexpected end of formula: expected a value, reference or function.');
      return missing();
    }
    switch (t.kind) {
      case 'number':
        next();
        return { type: 'Number', value: Number(t.text), raw: t.text, start: t.start, end: t.end };
      case 'string': {
        next();
        if (t.unterminated) error(t.start, t.end, 'syntax.unterminated-string', 'Unterminated text string: missing closing ".');
        const inner = t.text.slice(1, t.unterminated ? undefined : -1).replace(/""/g, '"');
        return { type: 'String', value: inner, raw: t.text, start: t.start, end: t.end };
      }
      case 'boolean':
        next();
        return { type: 'Boolean', value: t.text.toUpperCase() === 'TRUE', raw: t.text, start: t.start, end: t.end };
      case 'error':
        next();
        return { type: 'ErrorLiteral', raw: t.text, start: t.start, end: t.end };
      case 'cell':
      case 'columnRange':
      case 'rowRange':
      case 'name':
      case 'tableRef': {
        next();
        const body = t.text.slice(t.prefixLength);
        const kind: Reference['kind'] =
          t.kind === 'tableRef' ? 'table' : body === '#REF!' ? 'error' : t.kind === 'name' ? 'name' : t.kind;
        const ref: Reference = { type: 'Reference', kind, text: t.text, body, start: t.start, end: t.end };
        if (t.prefixLength > 0) ref.prefix = t.text.slice(0, t.prefixLength);
        return ref;
      }
      case 'function':
        return parseFunction();
      case 'lparen': {
        next();
        const saved = allowUnion;
        allowUnion = true;
        const expr = parseExpression(0);
        allowUnion = saved;
        let closed = false;
        let end = Math.max(expr.end, t.end);
        if (peek()?.kind === 'rparen') {
          end = next().end;
          closed = true;
        } else {
          const bad = peek();
          if (bad) {
            error(bad.start, bad.end, 'syntax.expected-paren', `Expected ')' but found ${describe(bad)}.`);
            skipTo(new Set(['rparen']));
            if (peek()?.kind === 'rparen') {
              end = next().end;
              closed = true;
            }
          }
          if (!closed) error(t.start, t.end, 'syntax.unclosed-paren', "Missing closing ')' for this parenthesis.");
        }
        return { type: 'Paren', expr, closed, start: t.start, end };
      }
      case 'lbrace':
        return parseArray();
      case 'rparen':
      case 'comma':
      case 'semicolon':
      case 'rbrace':
        error(t.start, t.end, 'syntax.expected-expression', `Expected a value, reference or function before ${describe(t)}.`);
        return missing();
      case 'colon':
        error(t.start, t.end, 'syntax.expected-expression', "Expected a reference before ':'.");
        return missing();
      case 'operator':
        if (t.text === '#') {
          next();
          error(t.start, t.end, 'syntax.bad-token', "Unexpected '#'. Error values look like #N/A, #REF!, #NAME?, #VALUE!, #DIV/0!.");
          return { type: 'Invalid', raw: t.text, start: t.start, end: t.end };
        }
        error(t.start, t.end, 'syntax.expected-expression', `Expected a value, reference or function before '${t.text}'.`);
        return missing();
      default:
        next();
        error(t.start, t.end, 'syntax.bad-token', `Unexpected character ${JSON.stringify(t.text)}.`);
        return { type: 'Invalid', raw: t.text, start: t.start, end: t.end };
    }
  }

  /** Parse `arg, arg, ...)` after an already-consumed `(`. */
  function parseArgs(open: Token, label: string): { args: Node[]; closed: boolean; end: number } {
    const saved = allowUnion;
    allowUnion = false;
    const args: Node[] = [];
    let closed = false;
    let end = open.end;

    if (peek()?.kind === 'rparen') {
      end = next().end;
      closed = true;
    } else {
      const stops = new Set(['comma', 'rparen']);
      for (;;) {
        const t = peek();
        if (t && (t.kind === 'comma' || t.kind === 'rparen')) args.push(missing());
        else args.push(parseExpression(0));
        let after = peek();
        if (after && after.kind !== 'comma' && after.kind !== 'rparen') {
          error(after.start, after.end, 'syntax.expected-comma', `Expected ',' or ')' but found ${describe(after)}.`);
          const junk = skipTo(stops);
          if (junk) args.push(junk);
          after = peek();
        }
        if (!after) {
          error(open.start, open.end, 'syntax.unclosed-paren', `Missing closing ')' for ${label}.`);
          end = args[args.length - 1]?.end ?? open.end;
          break;
        }
        end = next().end;
        if (after.kind === 'rparen') {
          closed = true;
          break;
        }
      }
    }
    allowUnion = saved;
    return { args, closed, end };
  }

  function parseFunction(): Node {
    const nameTok = next();
    const open = next(); // `(` is guaranteed by the tokenizer
    const { args, closed, end } = parseArgs(open, `${nameTok.text}(`);
    return {
      type: 'Function',
      name: nameTok.text,
      nameSpan: { start: nameTok.start, end: nameTok.end },
      args,
      closed,
      start: nameTok.start,
      end,
    };
  }

  function parseArray(): Node {
    const open = next();
    const saved = allowUnion;
    allowUnion = false;
    const rows: Node[][] = [[]];
    let closed = false;
    let end = open.end;
    const stops = new Set(['comma', 'semicolon', 'rbrace']);
    for (;;) {
      const t = peek();
      if (t && stops.has(t.kind)) {
        if (!(t.kind === 'rbrace' && rows.length === 1 && rows[0]!.length === 0)) {
          error(t.start, t.end, 'syntax.empty-array-element', 'Array constants cannot contain empty elements.');
        }
        rows[rows.length - 1]!.push(missing());
      } else {
        rows[rows.length - 1]!.push(parseExpression(0));
      }
      let after = peek();
      if (after && !stops.has(after.kind)) {
        error(after.start, after.end, 'syntax.expected-comma', `Expected ',', ';' or '}' but found ${describe(after)}.`);
        const junk = skipTo(stops);
        if (junk) rows[rows.length - 1]!.push(junk);
        after = peek();
      }
      if (!after) {
        error(open.start, open.end, 'syntax.unclosed-brace', "Missing closing '}' for this array constant.");
        end = rows[rows.length - 1]!.at(-1)?.end ?? open.end;
        break;
      }
      end = next().end;
      if (after.kind === 'rbrace') {
        closed = true;
        break;
      }
      if (after.kind === 'semicolon') rows.push([]);
    }
    allowUnion = saved;
    return { type: 'Array', rows, closed, start: open.start, end };
  }

  const body = parseExpression(0);

  // Anything left over is stray input.
  let reported = 0;
  while (peek()) {
    const t = next();
    if (reported++ < 5) {
      if (t.kind === 'rparen') error(t.start, t.end, 'syntax.unmatched-paren', "Unmatched ')': there is no opening parenthesis for it.");
      else if (t.kind === 'rbrace') error(t.start, t.end, 'syntax.unmatched-brace', "Unmatched '}'.");
      else if (t.kind === 'unknown') error(t.start, t.end, 'syntax.bad-token', `Unexpected character ${JSON.stringify(t.text)}.`);
      else error(t.start, t.end, 'syntax.unexpected-token', `Unexpected ${describe(t)}.`);
    }
  }

  // An empty non-formula is not an error; an empty formula `=` already reported "expected expression".
  const formula: Formula = { type: 'Formula', hasEquals, body, start: 0, end: src.length };
  if (!hasEquals && src.trim() === '') diagnostics.length = 0;
  return { formula, tokens, diagnostics };
}
