/** Source span, as offsets into the full formula text (including the leading `=`). */
export interface Span {
  start: number;
  end: number;
}

export type ReferenceKind = 'cell' | 'columnRange' | 'rowRange' | 'name' | 'table' | 'error';

export interface NumberLiteral extends Span {
  type: 'Number';
  value: number;
  raw: string;
}
export interface StringLiteral extends Span {
  type: 'String';
  value: string;
  raw: string;
}
export interface BooleanLiteral extends Span {
  type: 'Boolean';
  value: boolean;
  raw: string;
}
export interface ErrorLiteral extends Span {
  type: 'ErrorLiteral';
  raw: string;
}
export interface ArrayLiteral extends Span {
  type: 'Array';
  rows: Node[][];
  closed: boolean;
}
export interface Reference extends Span {
  type: 'Reference';
  kind: ReferenceKind;
  /** Full source text of the token, including any sheet prefix. */
  text: string;
  /** Sheet prefix including the trailing `!`, e.g. `'My Sheet'!`. */
  prefix?: string;
  /** Text after the prefix (the cell, range, name or table reference). */
  body: string;
}
export interface FunctionCall extends Span {
  type: 'Function';
  /** Name exactly as written. */
  name: string;
  nameSpan: Span;
  /** Arguments; an empty slot (`IF(a,,b)`) is a `Missing` node. */
  args: Node[];
  closed: boolean;
}
/** Immediate invocation of an expression: `LAMBDA(x, x+1)(5)`. */
export interface Call extends Span {
  type: 'Call';
  callee: Node;
  args: Node[];
  closed: boolean;
}
export interface Unary extends Span {
  type: 'Unary';
  op: '+' | '-' | '@';
  operand: Node;
}
export interface Postfix extends Span {
  type: 'Postfix';
  op: '%' | '#';
  operand: Node;
}
/** `:` range, ` ` intersection and `,` union are binary operators too. */
export interface Binary extends Span {
  type: 'Binary';
  op: string;
  opSpan: Span;
  left: Node;
  right: Node;
}
export interface Paren extends Span {
  type: 'Paren';
  expr: Node;
  closed: boolean;
}
/** An empty argument slot or a missing operand. */
export interface Missing extends Span {
  type: 'Missing';
}
/** A token the parser could not make sense of. */
export interface Invalid extends Span {
  type: 'Invalid';
  raw: string;
}

export type Node =
  | NumberLiteral
  | StringLiteral
  | BooleanLiteral
  | ErrorLiteral
  | ArrayLiteral
  | Reference
  | FunctionCall
  | Call
  | Unary
  | Postfix
  | Binary
  | Paren
  | Missing
  | Invalid;

export interface Formula extends Span {
  type: 'Formula';
  hasEquals: boolean;
  body: Node;
}

export type Severity = 'error' | 'warning' | 'info' | 'hint';

export interface Diagnostic extends Span {
  severity: Severity;
  code: string;
  message: string;
  /** Rendered faded in the editor (unused variables). */
  unnecessary?: boolean;
  /** Suggested replacement text for the span, offered as a quick fix. */
  fix?: { title: string; text: string };
}

export function children(node: Node): Node[] {
  switch (node.type) {
    case 'Array':
      return node.rows.flat();
    case 'Function':
      return node.args;
    case 'Call':
      return [node.callee, ...node.args];
    case 'Unary':
    case 'Postfix':
      return [node.operand];
    case 'Binary':
      return [node.left, node.right];
    case 'Paren':
      return [node.expr];
    default:
      return [];
  }
}

export function walk(node: Node, visit: (n: Node, parent: Node | undefined) => void, parent?: Node): void {
  visit(node, parent);
  for (const c of children(node)) walk(c, visit, node);
}
