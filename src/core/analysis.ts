import type { Diagnostic, FunctionCall, Node, Span } from './ast';
import { activeParamIndex, checkArity, getFunction, suggestFunctions } from './functions';
import { parse, type ParseResult } from './parser';
import { isOutOfGridCellLike } from './tokenizer';

/** Excel's documented hard limits. */
export const MAX_FORMULA_LENGTH = 8192;
export const MAX_NESTING = 64;

export interface LintContext {
  /**
   * Defined names in the workbook, lower-cased. When omitted, undefined-name checks are skipped
   * (e.g. when running outside Excel).
   */
  definedNames?: ReadonlySet<string>;
  /** Named LAMBDA functions in the workbook (lower-cased name -> accepted argument counts). */
  namedLambdas?: ReadonlyMap<string, { min: number; max: number }>;
}

/** A name introduced by LET or LAMBDA. */
export interface Sym {
  name: string;
  kind: 'variable' | 'parameter';
  declaration: Span;
  /** The symbol is visible for offsets strictly after scopeStart and up to scopeEnd. */
  scopeStart: number;
  scopeEnd: number;
  uses: Span[];
  /** The expression a LET variable is bound to. */
  value?: Node;
  optional?: boolean;
}

export interface Analysis {
  parse: ParseResult;
  diagnostics: Diagnostic[];
  symbols: Sym[];
}

type Scope = Map<string, Sym>;

const R1C1 = /^R(\[?-?\d+\]?)?C(\[?-?\d+\]?)?$/i;

export function analyze(src: string, ctx: LintContext = {}): Analysis {
  const result = parse(src);
  // Text without a leading `=` is a constant, not a formula: nothing to lint.
  const diagnostics: Diagnostic[] = result.formula.hasEquals ? [...result.diagnostics] : [];
  const symbols: Sym[] = [];
  /** Names that are really unreplaced snippet placeholders (e.g. `lookup_value` left in VLOOKUP's first slot). */
  const placeholders = new Set<Node>();
  const { formula } = result;

  const add = (d: Diagnostic) => diagnostics.push(d);

  if (formula.hasEquals && src.length > MAX_FORMULA_LENGTH) {
    add({
      severity: 'error',
      code: 'limit.formula-length',
      message: `Formula is ${src.length} characters long; Excel allows at most ${MAX_FORMULA_LENGTH}.`,
      start: MAX_FORMULA_LENGTH,
      end: src.length,
    });
  }

  function resolve(scopes: Scope[], name: string): Sym | undefined {
    const key = name.toLowerCase();
    for (let i = scopes.length - 1; i >= 0; i--) {
      const s = scopes[i]!.get(key);
      if (s) return s;
    }
    return undefined;
  }

  function declare(
    scope: Scope,
    node: Node,
    kind: Sym['kind'],
    scopeStart: number,
    scopeEnd: number,
    value?: Node,
  ): Sym | undefined {
    let name: string | undefined;
    let optional = false;
    if (node.type === 'Reference' && node.kind === 'name' && !node.prefix) name = node.text;
    else if (node.type === 'Reference' && node.kind === 'table' && kind === 'parameter' && /^\[[^\[\]@#]+\]$/.test(node.text)) {
      name = node.text.slice(1, -1);
      optional = true;
    }
    if (name === undefined) {
      if (node.type === 'Reference' && node.kind === 'cell') {
        add({
          severity: 'error',
          code: 'name.looks-like-cell',
          message: `'${node.text}' cannot be used as a name because it looks like a cell reference.`,
          start: node.start,
          end: node.end,
        });
      } else if (node.type !== 'Missing' && node.type !== 'Invalid') {
        add({
          severity: 'error',
          code: 'name.expected',
          message: kind === 'variable' ? 'Expected a variable name here.' : 'Expected a parameter name here.',
          start: node.start,
          end: node.end,
        });
      }
      return undefined;
    }
    const key = name.toLowerCase();
    const sym: Sym = {
      name,
      kind,
      declaration: { start: node.start, end: node.end },
      scopeStart,
      scopeEnd,
      uses: [],
      value,
      optional,
    };
    if (scope.has(key)) {
      add({
        severity: kind === 'parameter' ? 'error' : 'warning',
        code: 'name.duplicate',
        message:
          kind === 'parameter'
            ? `Parameter '${name}' is declared more than once.`
            : `'${name}' is already defined earlier in this LET; this definition replaces it from here on.`,
        start: node.start,
        end: node.end,
      });
    }
    scope.set(key, sym);
    symbols.push(sym);
    return sym;
  }

  function lambdaArity(sym: Sym | undefined): { min: number; max: number } | undefined {
    const v = sym?.value;
    if (v?.type !== 'Function' || v.name.toUpperCase() !== 'LAMBDA' || v.args.length === 0) return undefined;
    const params = v.args.slice(0, -1);
    const optional = params.filter((p) => p.type === 'Reference' && p.kind === 'table').length;
    return { min: params.length - optional, max: params.length };
  }

  function checkLambdaArity(min: number, max: number, argc: number, label: string, at: Span) {
    if (argc < min || argc > max) {
      add({
        severity: 'error',
        code: 'arity.lambda',
        message: `${label} expects ${min === max ? min : `${min} to ${max}`} argument${max === 1 ? '' : 's'}, but got ${argc}.`,
        start: at.start,
        end: at.end,
      });
    }
  }

  function visitFunction(fn: FunctionCall, scopes: Scope[], depth: number) {
    const upper = fn.name.toUpperCase();
    const nameSpan = fn.nameSpan;
    const scopeEnd = fn.closed ? fn.end : Number.POSITIVE_INFINITY;
    const hasInvalid = fn.args.some((a) => a.type === 'Invalid');

    if (depth > MAX_NESTING) {
      add({
        severity: 'error',
        code: 'limit.nesting',
        message: `Functions are nested more than ${MAX_NESTING} levels deep, which is Excel's limit.`,
        start: nameSpan.start,
        end: nameSpan.end,
      });
    }

    const info = getFunction(fn.name);
    const local = info ? undefined : resolve(scopes, fn.name);

    if (info) {
      const upperName = info.name;
      fn.args.forEach((arg, i) => {
        if (arg.type !== 'Reference' || arg.kind !== 'name' || arg.prefix) return;
        // LET names and LAMBDA parameters are declarations, not values.
        if (upperName === 'LET' && i % 2 === 0 && i < fn.args.length - 1) return;
        if (upperName === 'LAMBDA' && i < fn.args.length - 1) return;
        const param = info.params[activeParamIndex(info, i, fn.args.length)];
        if (!param || param.name.toLowerCase() !== arg.text.toLowerCase()) return;
        if (resolve(scopes, arg.text) || ctx.definedNames?.has(arg.text.toLowerCase())) return;
        placeholders.add(arg);
        add({
          severity: 'error',
          code: 'placeholder',
          message: `Replace the placeholder '${arg.text}' with a ${/range|array|table|ref/.test(param.name) ? 'range or value' : 'value or reference'}.`,
          start: arg.start,
          end: arg.end,
        });
      });
      if (!hasInvalid && fn.closed) {
        const msg = checkArity(info, fn.args.length);
        if (msg) {
          const tooMany = fn.args.length > info.maxArgs;
          const extra = fn.args[info.maxArgs];
          const closeParen = { start: fn.end - 1, end: fn.end };
          const at = tooMany && extra ? { start: extra.start, end: fn.args[fn.args.length - 1]!.end } : fn.args.length < info.minArgs ? closeParen : nameSpan;
          add({ severity: 'error', code: 'arity', message: msg, start: at.start, end: at.end });
        }
      }
    } else if (local) {
      local.uses.push(nameSpan);
      const ar = lambdaArity(local);
      if (ar && fn.closed && !hasInvalid) checkLambdaArity(ar.min, ar.max, fn.args.length, `'${local.name}'`, nameSpan);
    } else {
      const lower = fn.name.toLowerCase();
      const lambda = ctx.namedLambdas?.get(lower);
      if (lambda) {
        if (fn.closed && !hasInvalid) checkLambdaArity(lambda.min, lambda.max, fn.args.length, fn.name, nameSpan);
      } else if (!ctx.definedNames?.has(lower)) {
        const suggestions = suggestFunctions(fn.name);
        const hint = suggestions.length ? ` Did you mean ${suggestions.map((s) => `${s}`).join(', ')}?` : '';
        add({
          severity: 'warning',
          code: 'function.unknown',
          message: `Unknown function '${fn.name}': Excel will return #NAME? unless this is a defined name.${hint}`,
          start: nameSpan.start,
          end: nameSpan.end,
          fix: suggestions[0] ? { title: `Change to ${suggestions[0]}`, text: suggestions[0] } : undefined,
        });
      }
    }

    // LET: names are visible to later values and the final calculation.
    if (upper === 'LET') {
      const scope: Scope = new Map();
      const inner = [...scopes, scope];
      const a = fn.args;
      let i = 0;
      for (; i + 1 < a.length; i += 2) {
        visit(a[i + 1]!, inner, depth); // value sees only earlier names
        declare(scope, a[i]!, 'variable', a[i + 1]!.end, scopeEnd, a[i + 1]!);
      }
      if (i < a.length) visit(a[i]!, inner, depth);
      return;
    }

    // LAMBDA: parameters are visible only in the body.
    if (upper === 'LAMBDA' && fn.args.length >= 1) {
      const scope: Scope = new Map();
      const body = fn.args[fn.args.length - 1]!;
      const params = fn.args.slice(0, -1);
      const bodyStart = params.length ? params[params.length - 1]!.end : nameSpan.end;
      for (const p of params) declare(scope, p, 'parameter', bodyStart, scopeEnd);
      visit(body, [...scopes, scope], depth);
      return;
    }

    for (const arg of fn.args) visit(arg, scopes, depth);
  }

  function visit(node: Node, scopes: Scope[], depth: number): void {
    switch (node.type) {
      case 'Function':
        visitFunction(node, scopes, depth + 1);
        return;
      case 'Call': {
        visit(node.callee, scopes, depth);
        for (const a of node.args) visit(a, scopes, depth);
        const c = node.callee;
        if (c.type === 'Function' && c.name.toUpperCase() === 'LAMBDA' && c.args.length && node.closed) {
          const params = c.args.slice(0, -1);
          const optional = params.filter((p) => p.type === 'Reference' && p.kind === 'table').length;
          checkLambdaArity(params.length - optional, params.length, node.args.length, 'This LAMBDA', { start: c.nameSpan.start, end: c.nameSpan.end });
        }
        return;
      }
      case 'Reference': {
        if (node.kind === 'error') {
          add({
            severity: 'warning',
            code: 'ref.deleted',
            message: 'This reference points to a deleted cell, range or sheet (#REF!).',
            start: node.start,
            end: node.end,
          });
        } else if (node.kind === 'table' && /^\[\s*\]$/.test(node.body)) {
          // `Sales[]` is valid (the whole table), but brackets with no table name and nothing inside are not.
          add({
            severity: 'error',
            code: 'ref.empty-structured',
            message: 'Empty brackets: put a column name inside, like [@Units], or a table name in front, like Sales[].',
            start: node.start,
            end: node.end,
          });
        } else if (node.kind === 'name' && !node.prefix && !placeholders.has(node)) {
          const sym = resolve(scopes, node.text);
          if (sym) {
            sym.uses.push({ start: node.start, end: node.end });
          } else if (isOutOfGridCellLike(node.text)) {
            add({
              severity: 'warning',
              code: 'ref.out-of-grid',
              message: `'${node.text}' looks like a cell reference but is outside the worksheet grid (columns A to XFD, rows 1 to 1,048,576).`,
              start: node.start,
              end: node.end,
            });
          } else if (R1C1.test(node.text)) {
            add({
              severity: 'info',
              code: 'ref.r1c1',
              message: 'R1C1-style references are not supported here; use A1 notation.',
              start: node.start,
              end: node.end,
            });
          } else if (ctx.definedNames && !ctx.definedNames.has(node.text.toLowerCase())) {
            add({
              severity: 'warning',
              code: 'name.undefined',
              message: `'${node.text}' is not a defined name in this workbook or in this formula, so it will return #NAME?.`,
              start: node.start,
              end: node.end,
            });
          }
        }
        return;
      }
      case 'ErrorLiteral':
        if (node.raw.toUpperCase() === '#REF!') {
          add({
            severity: 'warning',
            code: 'ref.deleted',
            message: 'Formula contains #REF!, usually left by a deleted cell, range or sheet.',
            start: node.start,
            end: node.end,
          });
        }
        return;
      case 'Binary': {
        // Walk the left spine iteratively: `=A1+A2+...+A2000` is a left-deep tree thousands of nodes tall.
        const spine: Extract<Node, { type: 'Binary' }>[] = [];
        let cur: Node = node;
        while (cur.type === 'Binary') {
          spine.push(cur);
          cur = cur.left;
        }
        visit(cur, scopes, depth);
        for (let i = spine.length - 1; i >= 0; i--) {
          const b = spine[i]!;
          if (b.op === '/' && b.right.type === 'Number' && b.right.value === 0) {
            add({
              severity: 'warning',
              code: 'math.divide-by-zero',
              message: 'Dividing by zero returns #DIV/0!.',
              start: b.right.start,
              end: b.right.end,
            });
          }
          if (b.op === ' ') {
            add({
              severity: 'info',
              code: 'op.intersection',
              message: 'A space between two references is the intersection operator. If you meant to separate arguments, use a comma.',
              start: b.opSpan.start,
              end: b.opSpan.end,
            });
          }
          visit(b.right, scopes, depth);
        }
        return;
      }
      case 'Array': {
        const width = node.rows[0]?.length ?? 0;
        for (const row of node.rows) {
          if (row.length !== width) {
            add({
              severity: 'error',
              code: 'array.ragged',
              message: 'Every row of an array constant must have the same number of elements.',
              start: node.start,
              end: node.end,
            });
            break;
          }
        }
        for (const e of node.rows.flat()) {
          const ok =
            e.type === 'Number' ||
            e.type === 'String' ||
            e.type === 'Boolean' ||
            e.type === 'ErrorLiteral' ||
            e.type === 'Missing' ||
            e.type === 'Invalid' ||
            (e.type === 'Unary' && (e.op === '-' || e.op === '+') && e.operand.type === 'Number');
          if (!ok) {
            add({
              severity: 'error',
              code: 'array.not-constant',
              message: 'Array constants can only contain numbers, text, TRUE/FALSE and error values.',
              start: e.start,
              end: e.end,
            });
          }
        }
        return;
      }
      case 'Unary':
      case 'Postfix':
        visit(node.operand, scopes, depth);
        return;
      case 'Paren':
        visit(node.expr, scopes, depth);
        return;
      default:
        return;
    }
  }

  if (formula.hasEquals) {
    visit(formula.body, [], 0);
    for (const sym of symbols) {
      if (sym.uses.length === 0) {
        diagnostics.push({
          severity: 'hint',
          code: 'name.unused',
          message: `'${sym.name}' is declared but never used.`,
          start: sym.declaration.start,
          end: sym.declaration.end,
          unnecessary: true,
        });
      }
    }
  }

  diagnostics.sort((a, b) => a.start - b.start || a.end - b.end);
  return { parse: result, diagnostics, symbols };
}

/** Variables and parameters visible at `offset`, innermost declarations last. */
export function symbolsAt(analysis: Analysis, offset: number): Sym[] {
  const visible = analysis.symbols.filter((s) => offset > s.scopeStart && offset <= s.scopeEnd);
  // Keep only the innermost definition for each name.
  const byName = new Map<string, Sym>();
  for (const s of visible) byName.set(s.name.toLowerCase(), s);
  return [...byName.values()];
}
