import type { Node, Reference } from './ast';
import { getFunction } from './functions';
import { parse } from './parser';

export interface FormatOptions {
  /** Preferred maximum line width. */
  width: number;
  /** Spaces per indentation level. */
  indent: number;
  /** Upper-case recognised built-in function names (Excel does this on entry anyway). */
  uppercaseFunctions: boolean;
  /** Always expand LET onto multiple lines once it has this many name/value pairs. */
  expandLetAt: number;
}

export const DEFAULT_FORMAT_OPTIONS: FormatOptions = {
  width: 80,
  indent: 4,
  uppercaseFunctions: true,
  expandLetAt: 2,
};

export interface FormatResult {
  text: string;
  /** False when the formula has syntax errors and was returned untouched. */
  ok: boolean;
}

// ---------------------------------------------------------------------------
// A small Wadler/Prettier-style document model
// ---------------------------------------------------------------------------

type Doc =
  | string
  | Doc[]
  | { k: 'group'; d: Doc; hard: boolean }
  | { k: 'indent'; d: Doc }
  | { k: 'line' } // space when flat, newline when broken
  | { k: 'soft' } // nothing when flat, newline when broken
  | { k: 'hard' }; // always a newline

const line: Doc = { k: 'line' };
const soft: Doc = { k: 'soft' };
const hard: Doc = { k: 'hard' };
const indent = (d: Doc): Doc => ({ k: 'indent', d });

function hasHard(d: Doc): boolean {
  if (typeof d === 'string') return false;
  if (Array.isArray(d)) return d.some(hasHard);
  if (d.k === 'hard') return true;
  if (d.k === 'group') return d.hard;
  if (d.k === 'indent') return hasHard(d.d);
  return false;
}
const group = (d: Doc): Doc => ({ k: 'group', d, hard: hasHard(d) });

function join(sep: Doc, items: Doc[]): Doc[] {
  const out: Doc[] = [];
  items.forEach((it, i) => {
    if (i > 0) out.push(sep);
    out.push(it);
  });
  return out;
}

type Mode = 'flat' | 'break';
type Cmd = [level: number, mode: Mode, doc: Doc];

function fits(first: Cmd, rest: Cmd[], remaining: number): boolean {
  const stack: Cmd[] = [first];
  let restIdx = rest.length;
  while (remaining >= 0) {
    if (stack.length === 0) {
      if (restIdx === 0) return true;
      stack.push(rest[--restIdx]!);
      continue;
    }
    const [level, mode, d] = stack.pop()!;
    if (typeof d === 'string') {
      remaining -= d.length;
    } else if (Array.isArray(d)) {
      for (let i = d.length - 1; i >= 0; i--) stack.push([level, mode, d[i]!]);
    } else if (d.k === 'indent') {
      stack.push([level + 1, mode, d.d]);
    } else if (d.k === 'group') {
      if (mode === 'flat' && d.hard) return false;
      stack.push([level, d.hard ? 'break' : mode, d.d]);
    } else if (d.k === 'line') {
      if (mode === 'break') return true;
      remaining -= 1;
    } else if (d.k === 'soft') {
      if (mode === 'break') return true;
    } else if (d.k === 'hard') {
      return mode === 'break';
    }
  }
  return false;
}

function render(doc: Doc, width: number, unit: string): string {
  let out = '';
  let pos = 0;
  const stack: Cmd[] = [[0, 'break', doc]];
  const emit = (s: string) => {
    out += s;
    const nl = s.lastIndexOf('\n');
    pos = nl < 0 ? pos + s.length : s.length - nl - 1;
  };
  while (stack.length) {
    const [level, mode, d] = stack.pop()!;
    if (typeof d === 'string') emit(d);
    else if (Array.isArray(d)) for (let i = d.length - 1; i >= 0; i--) stack.push([level, mode, d[i]!]);
    else if (d.k === 'indent') stack.push([level + 1, mode, d.d]);
    else if (d.k === 'group') {
      if (mode === 'flat' && !d.hard) stack.push([level, 'flat', d.d]);
      else if (d.hard) stack.push([level, 'break', d.d]);
      else stack.push([level, fits([level, 'flat', d.d], stack, width - pos) ? 'flat' : 'break', d.d]);
    } else if (d.k === 'line') {
      if (mode === 'flat') emit(' ');
      else emit('\n' + unit.repeat(level));
    } else if (d.k === 'soft') {
      if (mode === 'break') emit('\n' + unit.repeat(level));
    } else emit('\n' + unit.repeat(level));
  }
  return out.replace(/[ \t]+\n/g, '\n');
}

// ---------------------------------------------------------------------------
// AST -> Doc
// ---------------------------------------------------------------------------

const BINARY_LEVEL: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

function referenceText(r: Reference): string {
  // Excel normalises cell and range references to upper case; names keep their case.
  if (r.kind === 'cell' || r.kind === 'columnRange' || r.kind === 'rowRange') return (r.prefix ?? '') + r.body.toUpperCase();
  return r.text;
}

function functionName(name: string, opts: FormatOptions): string {
  return opts.uppercaseFunctions && getFunction(name) ? name.toUpperCase() : name;
}

function toDoc(n: Node, opts: FormatOptions): Doc {
  switch (n.type) {
    case 'Number':
    case 'String':
    case 'ErrorLiteral':
      return n.raw;
    case 'Boolean':
      return n.raw.toUpperCase();
    case 'Reference':
      return referenceText(n);
    case 'Missing':
      return '';
    case 'Invalid':
      return n.raw;
    case 'Unary':
      return [n.op, toDoc(n.operand, opts)];
    case 'Postfix':
      return [toDoc(n.operand, opts), n.op];
    case 'Paren':
      return group(['(', indent([soft, toDoc(n.expr, opts)]), soft, ')']);
    case 'Array':
      return group([
        '{',
        indent([soft, join([';', line], n.rows.map((row) => join(', ', row.map((e) => toDoc(e, opts)))))]),
        soft,
        '}',
      ]);
    case 'Function':
      return functionDoc(n, opts);
    case 'Call': {
      const args = n.args.map((a) => toDoc(a, opts));
      if (args.length === 0) return [toDoc(n.callee, opts), '()'];
      return [toDoc(n.callee, opts), group(['(', indent([soft, join([',', line], args)]), soft, ')'])];
    }
    case 'Binary': {
      if (n.op === ':') return [toDoc(n.left, opts), ':', toDoc(n.right, opts)];
      if (n.op === ' ') return [toDoc(n.left, opts), ' ', toDoc(n.right, opts)];
      if (n.op === ',') return [toDoc(n.left, opts), ',', toDoc(n.right, opts)];
      // Flatten a left-associative chain of operators at the same precedence level.
      const level = BINARY_LEVEL[n.op];
      const operands: Node[] = [];
      const ops: string[] = [];
      let cur: Node = n;
      while (cur.type === 'Binary' && BINARY_LEVEL[cur.op] === level) {
        operands.unshift(cur.right);
        ops.unshift(cur.op);
        cur = cur.left;
      }
      operands.unshift(cur);
      const rest: Doc[] = ops.map((op, i) => [line, op, ' ', toDoc(operands[i + 1]!, opts)]);
      return group([toDoc(operands[0]!, opts), indent(rest)]);
    }
  }
}

function functionDoc(n: Extract<Node, { type: 'Function' }>, opts: FormatOptions): Doc {
  const name = functionName(n.name, opts);
  const upper = n.name.toUpperCase();
  const args = n.args.map((a) => toDoc(a, opts));
  if (args.length === 0) return `${name}()`;

  if (upper === 'LET' && args.length >= 3 && args.length % 2 === 1) {
    const pairs: Doc[] = [];
    for (let i = 0; i + 1 < args.length; i += 2) pairs.push([args[i]!, ', ', args[i + 1]!]);
    const items = [...pairs, args[args.length - 1]!];
    const expand = pairs.length >= opts.expandLetAt;
    const sep: Doc = expand ? [',', hard] : [',', line];
    return group([`${name}(`, indent([expand ? hard : soft, join(sep, items)]), expand ? hard : soft, ')']);
  }

  if (upper === 'LAMBDA' && args.length >= 2) {
    const params = join(', ', args.slice(0, -1));
    return group([`${name}(`, indent([soft, params, ',', line, args[args.length - 1]!]), soft, ')']);
  }

  return group([`${name}(`, indent([soft, join([',', line], args)]), soft, ')']);
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export function formatFormula(src: string, options: Partial<FormatOptions> = {}): FormatResult {
  const opts = { ...DEFAULT_FORMAT_OPTIONS, ...options };
  const { formula, diagnostics } = parse(src);
  if (!formula.hasEquals || diagnostics.length > 0) return { text: src, ok: false };
  const doc: Doc = ['=', toDoc(formula.body, opts)];
  return { text: render(doc, opts.width, ' '.repeat(opts.indent)), ok: true };
}

function minifyNode(n: Node, opts: FormatOptions): string {
  switch (n.type) {
    case 'Number':
    case 'String':
    case 'ErrorLiteral':
      return n.raw;
    case 'Boolean':
      return n.raw.toUpperCase();
    case 'Reference':
      return referenceText(n);
    case 'Missing':
      return '';
    case 'Invalid':
      return n.raw;
    case 'Unary':
      return n.op + minifyNode(n.operand, opts);
    case 'Postfix':
      return minifyNode(n.operand, opts) + n.op;
    case 'Paren':
      return `(${minifyNode(n.expr, opts)})`;
    case 'Array':
      return `{${n.rows.map((r) => r.map((e) => minifyNode(e, opts)).join(',')).join(';')}}`;
    case 'Function':
      return `${functionName(n.name, opts)}(${n.args.map((a) => minifyNode(a, opts)).join(',')})`;
    case 'Call':
      return `${minifyNode(n.callee, opts)}(${n.args.map((a) => minifyNode(a, opts)).join(',')})`;
    case 'Binary': {
      // Iterate down the left spine so very long operator chains cannot overflow the stack.
      const spine: Extract<Node, { type: 'Binary' }>[] = [];
      let cur: Node = n;
      while (cur.type === 'Binary') {
        spine.push(cur);
        cur = cur.left;
      }
      let out = minifyNode(cur, opts);
      for (let i = spine.length - 1; i >= 0; i--) out += spine[i]!.op + minifyNode(spine[i]!.right, opts);
      return out;
    }
  }
}

/** Remove all insignificant whitespace (keeping the intersection operator's space). */
export function minifyFormula(src: string, options: Partial<FormatOptions> = {}): FormatResult {
  const opts = { ...DEFAULT_FORMAT_OPTIONS, ...options };
  const { formula, diagnostics } = parse(src);
  if (!formula.hasEquals || diagnostics.length > 0) return { text: src, ok: false };
  return { text: '=' + minifyNode(formula.body, opts), ok: true };
}
