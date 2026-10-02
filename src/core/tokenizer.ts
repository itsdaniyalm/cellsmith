export type TokenKind =
  | 'whitespace'
  | 'number'
  | 'string'
  | 'error'
  | 'boolean'
  | 'function' // identifier immediately followed by `(`
  | 'name' // identifier that is not a cell reference
  | 'cell' // A1, $A$1
  | 'columnRange' // A:A, $A:$C
  | 'rowRange' // 1:1, $2:$5
  | 'tableRef' // Table1[Col], [@Col], Table1[[#This Row],[Col]]
  | 'operator' // + - * / ^ & = <> < > <= >= % @ #
  | 'colon'
  | 'comma'
  | 'semicolon'
  | 'lparen'
  | 'rparen'
  | 'lbrace'
  | 'rbrace'
  | 'unknown';

export interface Token {
  kind: TokenKind;
  text: string;
  start: number;
  end: number;
  /** For reference-like tokens: length of the leading `Sheet!` prefix (0 if none). */
  prefixLength: number;
  /** Set on strings that run to end of input without a closing quote. */
  unterminated?: boolean;
}

export const MAX_COLUMN = 16384; // XFD
export const MAX_ROW = 1048576;

export const ERROR_LITERALS = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#SPILL!',
  '#CALC!',
  '#GETTING_DATA',
  '#FIELD!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#CONNECT!',
  '#EXTERNAL!',
];

const IDENT_START = /[\p{L}_\\]/u;
const IDENT_PART = /[\p{L}\p{N}_.\\]/u;
const IDENT = /[\p{L}_\\][\p{L}\p{N}_.\\]*/uy;
const CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
const CELL_TOKEN = /\$?[A-Za-z]{1,3}\$?\d+/y;
const COL_RANGE = /\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/y;
const ROW_RANGE = /\$?\d+:\$?\d+/y;
const NUMBER = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const WHITESPACE = /[ \t\r\n ]+/y;

/** Column letters -> 1-based index ("A" = 1, "XFD" = 16384). */
export function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** True when text is a syntactically valid in-grid cell reference such as `$B$12`. */
export function isCellReference(text: string): boolean {
  const m = CELL.exec(text);
  if (!m) return false;
  const row = Number(m[2]);
  return columnIndex(m[1]!) <= MAX_COLUMN && row >= 1 && row <= MAX_ROW;
}

/** True when text has the shape of a cell reference but falls outside the grid (e.g. `XFE1`, `A0`). */
export function isOutOfGridCellLike(text: string): boolean {
  return CELL.test(text) && !isCellReference(text);
}

function matchAt(re: RegExp, src: string, i: number): string | undefined {
  re.lastIndex = i;
  const m = re.exec(src);
  return m ? m[0] : undefined;
}

/**
 * Length of a `Sheet!` prefix starting at i (including the `!`), or 0.
 * Handles 'Quoted Name'!, Unquoted!, Sheet1:Sheet3! (3D), [Book]Sheet! and '[Book.xlsx]Sheet 1'!.
 */
function sheetPrefixLength(src: string, i: number): number {
  let j = i;
  const part = (): boolean => {
    if (src[j] === "'") {
      let k = j + 1;
      for (;;) {
        if (k >= src.length) return false;
        if (src[k] === "'") {
          if (src[k + 1] === "'") {
            k += 2;
            continue;
          }
          break;
        }
        k++;
      }
      j = k + 1;
      return true;
    }
    if (src[j] === '[') {
      const close = src.indexOf(']', j);
      if (close < 0) return false;
      j = close + 1;
    }
    const id = matchAt(IDENT, src, j);
    if (!id) return false;
    j += id.length;
    return true;
  };
  if (!part()) return 0;
  if (src[j] === ':') {
    const save = j;
    j++;
    if (!part()) j = save;
  }
  return src[j] === '!' ? j + 1 - i : 0;
}

/** Scan a bracketed table specifier starting at `[`; returns the end offset or -1 if unbalanced. */
function scanBrackets(src: string, i: number): number {
  let depth = 0;
  for (let j = i; j < src.length; j++) {
    const ch = src[j];
    if (ch === "'") {
      j++; // escaped character inside a table specifier
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') {
      depth--;
      if (depth === 0) return j + 1;
    }
  }
  return -1;
}

const OPERAND_ENDERS = new Set<TokenKind>([
  'number',
  'string',
  'error',
  'boolean',
  'name',
  'cell',
  'columnRange',
  'rowRange',
  'tableRef',
  'rparen',
  'rbrace',
]);

export function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  let prevSig: Token | undefined;

  const push = (kind: TokenKind, start: number, end: number, prefixLength = 0, extra?: Partial<Token>) => {
    const t: Token = { kind, text: src.slice(start, end), start, end, prefixLength, ...extra };
    tokens.push(t);
    if (kind !== 'whitespace') prevSig = t;
    i = end;
  };
  const afterOperand = () =>
    !!prevSig && (OPERAND_ENDERS.has(prevSig.kind) || (prevSig.kind === 'operator' && (prevSig.text === '%' || prevSig.text === '#')));

  /** Reference body (after an optional sheet prefix) starting at `p`. Returns false if nothing matched. */
  const scanReferenceBody = (start: number, p: number, prefixLength: number): boolean => {
    if (src.startsWith('#REF!', p)) {
      push('name', start, p + 5, prefixLength);
      return true;
    }
    const col = matchAt(COL_RANGE, src, p);
    if (col && !IDENT_PART.test(src[p + col.length] ?? '') && src[p + col.length] !== '(' && src[p + col.length] !== '[') {
      push('columnRange', start, p + col.length, prefixLength);
      return true;
    }
    const row = matchAt(ROW_RANGE, src, p);
    if (row && !IDENT_PART.test(src[p + row.length] ?? '')) {
      push('rowRange', start, p + row.length, prefixLength);
      return true;
    }
    const cell = matchAt(CELL_TOKEN, src, p);
    if (cell && isCellReference(cell)) {
      const after = src[p + cell.length] ?? '';
      if (!IDENT_PART.test(after) && after !== '(' && after !== '[') {
        push('cell', start, p + cell.length, prefixLength);
        return true;
      }
    }
    if (src[p] === '$') return false;
    const id = matchAt(IDENT, src, p);
    if (!id) return false;
    const end = p + id.length;
    const next = src[end];
    if (next === '[') {
      const close = scanBrackets(src, end);
      if (close > 0) {
        push('tableRef', start, close, prefixLength);
        return true;
      }
    }
    if (next === '(' && prefixLength === 0) {
      push('function', start, end, 0);
      return true;
    }
    push('name', start, end, prefixLength);
    return true;
  };

  while (i < src.length) {
    const ch = src[i]!;

    const ws = matchAt(WHITESPACE, src, i);
    if (ws) {
      push('whitespace', i, i + ws.length);
      continue;
    }

    if (ch === '"') {
      let j = i + 1;
      let closed = false;
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            j += 2;
            continue;
          }
          closed = true;
          j++;
          break;
        }
        j++;
      }
      push('string', i, j, 0, closed ? undefined : { unterminated: true });
      continue;
    }

    if (ch === '#') {
      // Spill operator: must directly follow the reference with no whitespace (A1#).
      if (afterOperand() && prevSig!.end === i) {
        push('operator', i, i + 1); // spill operator: A1#
        continue;
      }
      const upper = src.slice(i).toUpperCase();
      const lit = ERROR_LITERALS.find((e) => upper.startsWith(e));
      if (lit) {
        push('error', i, i + lit.length);
        continue;
      }
      push('operator', i, i + 1);
      continue;
    }

    // Row ranges (1:1) are only valid where an operand is expected, otherwise `1:2` could be junk after a value.
    if (/[\d$]/.test(ch) && !afterOperand()) {
      const row = matchAt(ROW_RANGE, src, i);
      if (row && !IDENT_PART.test(src[i + row.length] ?? '')) {
        push('rowRange', i, i + row.length);
        continue;
      }
    }

    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1] ?? ''))) {
      const num = matchAt(NUMBER, src, i)!;
      push('number', i, i + num.length);
      continue;
    }

    if (ch === "'" || ch === '[' || ch === '$' || IDENT_START.test(ch)) {
      const prefix = ch === '$' ? 0 : sheetPrefixLength(src, i);
      if (prefix > 0) {
        if (scanReferenceBody(i, i + prefix, prefix)) continue;
        // `Sheet1!` followed by nothing useful: emit prefix alone as an unknown token
        push('unknown', i, i + prefix);
        continue;
      }
      if (ch === '[') {
        const close = scanBrackets(src, i);
        if (close > 0) {
          push('tableRef', i, close);
          continue;
        }
        push('unknown', i, i + 1);
        continue;
      }
      if (ch === "'") {
        push('unknown', i, i + 1);
        continue;
      }
      if (ch === '$') {
        if (scanReferenceBody(i, i, 0)) continue;
        push('unknown', i, i + 1);
        continue;
      }
      // Plain identifier: function, boolean, cell, name or table reference
      const id = matchAt(IDENT, src, i)!;
      const end = i + id.length;
      const upper = id.toUpperCase();
      if ((upper === 'TRUE' || upper === 'FALSE') && src[end] !== '(') {
        push('boolean', i, end);
        continue;
      }
      if (scanReferenceBody(i, i, 0)) continue;
      push('name', i, end);
      continue;
    }

    switch (ch) {
      case '(':
        push('lparen', i, i + 1);
        continue;
      case ')':
        push('rparen', i, i + 1);
        continue;
      case '{':
        push('lbrace', i, i + 1);
        continue;
      case '}':
        push('rbrace', i, i + 1);
        continue;
      case ',':
        push('comma', i, i + 1);
        continue;
      case ';':
        push('semicolon', i, i + 1);
        continue;
      case ':':
        push('colon', i, i + 1);
        continue;
      case '<':
        push('operator', i, src[i + 1] === '>' || src[i + 1] === '=' ? i + 2 : i + 1);
        continue;
      case '>':
        push('operator', i, src[i + 1] === '=' ? i + 2 : i + 1);
        continue;
      case '+':
      case '-':
      case '*':
      case '/':
      case '^':
      case '&':
      case '=':
      case '%':
      case '@':
        push('operator', i, i + 1);
        continue;
      default:
        push('unknown', i, i + 1);
    }
  }
  return tokens;
}
