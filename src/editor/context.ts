import type { LintContext } from '../core/analysis';
import { parse } from '../core/parser';

export interface WorkbookTable {
  name: string;
  columns: string[];
}
export interface WorkbookName {
  name: string;
  /** The name's formula as Excel reports it, e.g. `=Sheet1!$A$1:$A$10` or `=LAMBDA(x, x+1)`. */
  refersTo: string;
}
export interface WorkbookInfo {
  sheets: string[];
  names: WorkbookName[];
  tables: WorkbookTable[];
}

export const EMPTY_WORKBOOK: WorkbookInfo = { sheets: [], names: [], tables: [] };

export interface NamedLambda {
  name: string;
  params: { name: string; optional: boolean }[];
  refersTo: string;
}

/** Extract workbook names whose definition is a LAMBDA, so they can be completed and arity-checked. */
export function namedLambdas(info: WorkbookInfo): NamedLambda[] {
  const out: NamedLambda[] = [];
  for (const n of info.names) {
    if (!/^=\s*LAMBDA\s*\(/i.test(n.refersTo)) continue;
    const body = parse(n.refersTo).formula.body;
    if (body.type !== 'Function' || body.args.length === 0) continue;
    const params = body.args.slice(0, -1).map((a) => {
      if (a.type === 'Reference' && a.kind === 'table') return { name: a.text.slice(1, -1), optional: true };
      return { name: a.type === 'Reference' ? a.text : '?', optional: false };
    });
    out.push({ name: n.name, params, refersTo: n.refersTo });
  }
  return out;
}

/**
 * Build the linter's view of the workbook. Returns undefined when we are not connected to a
 * workbook, which disables the "undefined name" checks rather than flagging every name.
 */
export function lintContextFor(info: WorkbookInfo | undefined): LintContext | undefined {
  if (!info) return undefined;
  const definedNames = new Set<string>();
  for (const n of info.names) definedNames.add(n.name.toLowerCase());
  for (const t of info.tables) definedNames.add(t.name.toLowerCase());
  const lambdas = new Map<string, { min: number; max: number }>();
  for (const l of namedLambdas(info)) {
    const optional = l.params.filter((p) => p.optional).length;
    lambdas.set(l.name.toLowerCase(), { min: l.params.length - optional, max: l.params.length });
  }
  return { definedNames, namedLambdas: lambdas };
}

/** Sheet names need quoting unless they are plain identifiers that cannot be mistaken for a cell. */
export function quoteSheetName(name: string): string {
  const plain = /^[A-Za-z_][\w.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name);
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Escape characters that have special meaning inside a table specifier. */
export function escapeTableItem(name: string): string {
  return name.replace(/(['\[\]#])/g, "'$1");
}
