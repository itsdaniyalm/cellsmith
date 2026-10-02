import { CATALOG } from './functionData';

export interface FunctionParam {
  name: string;
  optional: boolean;
}

export interface FunctionInfo {
  /** Canonical upper-case name, e.g. `STDEV.S`. */
  name: string;
  category: string;
  description: string;
  /** Display signature, e.g. `IF(logical_test, [value_if_true], [value_if_false])`. */
  signature: string;
  params: FunctionParam[];
  /** Size of the trailing group of parameters that may repeat (0 = none). */
  repeat: number;
  minArgs: number;
  maxArgs: number;
  volatile: boolean;
  /** Extra argument-count rule; returns a message when the count is invalid. */
  arityRule?: (argc: number) => string | undefined;
  /** Overrides which parameter is highlighted in signature help. */
  activeParam?: (argIndex: number, argc: number) => number;
}

/** Excel allows at most 255 arguments in a function call. */
export const MAX_ARGS = 255;

const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT', 'INFO']);

/** Split on commas that are not nested inside [] brackets. */
function splitTopLevel(s: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let cur = '';
  for (const ch of s) {
    if (ch === '[') depth++;
    if (ch === ']') depth--;
    if (ch === ',' && depth === 0) {
      out.push(cur.trim());
      cur = '';
    } else cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}

function buildInfo(signature: string, description: string, category: string): FunctionInfo {
  const open = signature.indexOf('(');
  const name = signature.slice(0, open);
  const inner = signature.slice(open + 1, signature.lastIndexOf(')'));
  const params: FunctionParam[] = [];
  let lastGroup = 0;
  let repeat = 0;
  for (const item of splitTopLevel(inner)) {
    if (item === '...') {
      repeat = lastGroup;
    } else if (item.startsWith('[')) {
      const names = item.slice(1, -1).split(',').map((s) => s.trim());
      for (const n of names) params.push({ name: n, optional: true });
      lastGroup = names.length;
    } else {
      params.push({ name: item, optional: false });
      lastGroup = 1;
    }
  }
  const required = params.filter((p) => !p.optional).length;
  return {
    name,
    category,
    description,
    signature,
    params,
    repeat,
    minArgs: required,
    maxArgs: repeat ? MAX_ARGS : params.length,
    volatile: VOLATILE.has(name),
  };
}

const registry = new Map<string, FunctionInfo>();
for (const [category, entries] of Object.entries(CATALOG)) {
  for (const [signature, description] of entries) {
    const info = buildInfo(signature, description, category);
    registry.set(info.name, info);
  }
}

const parityRule = (parity: 'odd' | 'even', what: string) => (argc: number) =>
  argc % 2 === (parity === 'odd' ? 1 : 0) ? undefined : `${what} (got ${argc} argument${argc === 1 ? '' : 's'}).`;

function override(name: string, patch: Partial<FunctionInfo>): void {
  const info = registry.get(name);
  if (info) Object.assign(info, patch);
}

override('LET', {
  minArgs: 3,
  arityRule: parityRule('odd', 'LET needs name/value pairs followed by a final calculation: an odd number of arguments'),
});
override('IFS', { arityRule: parityRule('even', 'IFS needs condition/value pairs: an even number of arguments') });
override('COUNTIFS', { arityRule: parityRule('even', 'COUNTIFS needs range/criteria pairs: an even number of arguments') });
for (const n of ['SUMIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS']) {
  override(n, { arityRule: parityRule('odd', `${n} needs a range followed by range/criteria pairs: an odd number of arguments`) });
}
override('LAMBDA', {
  params: [
    { name: 'parameter1', optional: true },
    { name: 'calculation', optional: false },
  ],
  signature: 'LAMBDA([parameter1, parameter2, ...], calculation)',
  repeat: 0,
  minArgs: 1,
  maxArgs: MAX_ARGS,
  activeParam: (i, argc) => (argc > 1 && i === argc - 1 ? 1 : 0),
});
override('MAP', {
  signature: 'MAP(array1, [array2, ...], lambda)',
  params: [
    { name: 'array1', optional: false },
    { name: 'array2', optional: true },
    { name: 'lambda', optional: false },
  ],
  repeat: 0,
  minArgs: 2,
  activeParam: (i, argc) => (i === argc - 1 && argc > 1 ? 2 : Math.min(i, 1)),
});

/** Strip the internal prefixes Excel stores for newer functions (`_xlfn.`, `_xlfn._xlws.`, `_xlpm.`). */
export function normalizeFunctionName(name: string): string {
  return name.replace(/^(?:_xlfn\.)?(?:_xlws\.)?/i, '').replace(/^_xlpm\./i, '').toUpperCase();
}

export function getFunction(name: string): FunctionInfo | undefined {
  return registry.get(normalizeFunctionName(name));
}

export function allFunctions(): FunctionInfo[] {
  return [...registry.values()];
}

/** Which parameter should be highlighted when the caret is in argument `argIndex` of `argc`. */
export function activeParamIndex(info: FunctionInfo, argIndex: number, argc: number): number {
  if (info.activeParam) return info.activeParam(argIndex, argc);
  const n = info.params.length;
  if (argIndex < n) return argIndex;
  if (info.repeat > 0) return n - info.repeat + ((argIndex - n) % info.repeat);
  return n - 1;
}

/** Why `argc` arguments is wrong for this function, or undefined if acceptable. */
export function checkArity(info: FunctionInfo, argc: number): string | undefined {
  if (argc < info.minArgs || argc > info.maxArgs) {
    const range =
      info.maxArgs === MAX_ARGS && info.repeat
        ? `at least ${info.minArgs}`
        : info.minArgs === info.maxArgs
          ? `${info.minArgs}`
          : `${info.minArgs} to ${info.maxArgs}`;
    const got = `${argc} argument${argc === 1 ? '' : 's'}`;
    if (info.maxArgs === 0) return `${info.name} takes no arguments (got ${got}).`;
    if (argc > MAX_ARGS) return `Too many arguments: Excel allows at most ${MAX_ARGS}.`;
    return `${info.name} expects ${range} argument${range === '1' ? '' : 's'}, but got ${got}.`;
  }
  return info.arityRule?.(argc);
}

/** Closest known function names to `name` (for "did you mean" suggestions). */
export function suggestFunctions(name: string, limit = 3): string[] {
  const target = normalizeFunctionName(name);
  const scored: [string, number][] = [];
  for (const fname of registry.keys()) {
    if (fname === target || Math.abs(fname.length - target.length) > 2) continue;
    const d = editDistance(target, fname, 2);
    if (d <= 2 && d < Math.max(target.length, 3) - 1) scored.push([fname, d]);
  }
  return scored.sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0])).slice(0, limit).map(([n]) => n);
}

function editDistance(a: string, b: string, max: number): number {
  const m = a.length;
  const n = b.length;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    let rowMin = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      const v = Math.min(prev[j]! + 1, cur[j - 1]! + 1, prev[j - 1]! + cost);
      cur[j] = v;
      if (v < rowMin) rowMin = v;
    }
    if (rowMin > max) return max + 1;
    prev = cur;
  }
  return prev[n]!;
}
