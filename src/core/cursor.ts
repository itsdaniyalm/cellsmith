import { children, type FunctionCall, type Node } from './ast';
import type { ParseResult } from './parser';

export type CompletionContext =
  | { kind: 'none' }
  /** Inside a table specifier such as `Table1[|` or `Table1[[#This Row],[|`. */
  | { kind: 'table'; table?: string; start: number; end: number }
  /** Typing an identifier; `start`..`end` is the text a completion replaces. */
  | { kind: 'word'; word: string; start: number; end: number; nextChar: string };

const WORD_BEFORE = /[A-Za-z_\\][\w.\\]*$/;
const WORD_AFTER = /^[\w.\\]*/;

/** True when `offset` falls inside an open (or just-closed-at-cursor) string literal. */
function insideString(text: string, offset: number): boolean {
  let quotes = 0;
  for (let i = 0; i < offset; i++) if (text[i] === '"') quotes++;
  return quotes % 2 === 1;
}

/** Where the caret sits relative to table brackets: the start offsets of all open `[`, outermost first. */
function openBrackets(text: string, offset: number): number[] {
  const stack: number[] = [];
  let inString = false;
  for (let i = 0; i < offset; i++) {
    const ch = text[i];
    if (ch === '"' && stack.length === 0) inString = !inString;
    if (inString) continue;
    if (stack.length > 0 && ch === "'") {
      i++; // escaped character inside a table specifier
      continue;
    }
    if (ch === '[') stack.push(i);
    else if (ch === ']') stack.pop();
  }
  return stack;
}

export function completionContext(text: string, offset: number): CompletionContext {
  if (!text.startsWith('=') || offset < 1 || insideString(text, offset)) return { kind: 'none' };

  const open = openBrackets(text, offset);
  if (open.length > 0) {
    const inner = open[open.length - 1]!;
    const before = text.slice(0, open[0]!);
    const table = /([A-Za-z_À-￿][\w.À-￿]*)$/.exec(before)?.[1];
    let start = inner + 1;
    if (text[start] === '@' && offset > start) start++; // `[@Col` - complete the column after the @
    return { kind: 'table', table, start, end: offset };
  }

  const before = text.slice(0, offset);
  const m = WORD_BEFORE.exec(before);
  const word = m?.[0] ?? '';
  const start = offset - word.length;
  // A word directly after `Sheet!` is a cell or name on that sheet: nothing useful to suggest.
  if (text[start - 1] === '!') return { kind: 'none' };
  // Cell references like $B$2 or digits glued to a word are not completable identifiers.
  if (text[start - 1] === '$') return { kind: 'none' };
  const after = WORD_AFTER.exec(text.slice(offset))![0];
  const end = offset + after.length;
  return { kind: 'word', word, start, end, nextChar: text[end] ?? '' };
}

export interface CallContext {
  fn: FunctionCall;
  /** Zero-based index of the argument the caret is in. */
  argIndex: number;
  /** Number of arguments typed so far, counting the one under the caret. */
  argc: number;
}

/** The innermost function call whose argument list contains `offset`. */
export function enclosingCall(parsed: ParseResult, offset: number): CallContext | undefined {
  let found: FunctionCall | undefined;
  const visit = (n: Node) => {
    if (n.type === 'Function' && offset > n.nameSpan.end && (n.closed ? offset < n.end : true)) found = n;
    for (const c of children(n)) visit(c);
  };
  visit(parsed.formula.body);
  if (!found) return undefined;

  const openParen = found.nameSpan.end;
  let depth = 0;
  let argIndex = 0;
  for (const t of parsed.tokens) {
    if (t.start <= openParen) continue;
    if (t.end > offset) break;
    if (t.kind === 'lparen' || t.kind === 'lbrace') depth++;
    else if (t.kind === 'rparen' || t.kind === 'rbrace') depth = Math.max(0, depth - 1);
    else if (t.kind === 'comma' && depth === 0) argIndex++;
  }
  return { fn: found, argIndex, argc: Math.max(found.args.length, argIndex + 1) };
}

/** The chain of nodes from the root down to the deepest node containing `offset`. */
export function nodePathAt(root: Node, offset: number): Node[] {
  const path: Node[] = [];
  let cur: Node | undefined = root;
  while (cur) {
    path.push(cur);
    cur = children(cur).find((c) => c.start <= offset && offset <= c.end && c.end > c.start);
  }
  return path;
}
