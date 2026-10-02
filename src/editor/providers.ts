import type * as M from 'monaco-editor';
import { analyze, symbolsAt, type Analysis, type Sym } from '../core/analysis';
import { completionContext, enclosingCall, nodePathAt } from '../core/cursor';
import { formatFormula, type FormatOptions } from '../core/formatter';
import { activeParamIndex, allFunctions, getFunction, type FunctionInfo } from '../core/functions';
import { escapeTableItem, lintContextFor, namedLambdas, quoteSheetName, type WorkbookInfo } from './context';
import { LANGUAGE_ID } from './language';
import { monaco } from './monaco';

/** What the providers need to know about the host app. */
export interface EditorServices {
  /** Live workbook metadata, or undefined when not connected to Excel. */
  getWorkbook(): WorkbookInfo | undefined;
  /** Bumped whenever the workbook metadata changes, so cached analyses are recomputed. */
  workbookVersion(): number;
  getFormatOptions(): Partial<FormatOptions>;
}

const MARKER_OWNER = 'cellsmith';

// ---------------------------------------------------------------------------
// Analysis cache (one per model version)
// ---------------------------------------------------------------------------

const cache = new WeakMap<M.editor.ITextModel, { version: number; wb: number; analysis: Analysis }>();

export function getAnalysis(model: M.editor.ITextModel, svc: EditorServices): Analysis {
  const version = model.getVersionId();
  const wb = svc.workbookVersion();
  const hit = cache.get(model);
  if (hit && hit.version === version && hit.wb === wb) return hit.analysis;
  const analysis = analyze(model.getValue(), lintContextFor(svc.getWorkbook()));
  cache.set(model, { version, wb, analysis });
  return analysis;
}

const SEVERITY = {
  error: monaco.MarkerSeverity.Error,
  warning: monaco.MarkerSeverity.Warning,
  info: monaco.MarkerSeverity.Info,
  hint: monaco.MarkerSeverity.Hint,
} as const;

/** Run the analyzer and publish its diagnostics as editor markers. */
export function validateModel(model: M.editor.ITextModel, svc: EditorServices): Analysis {
  const analysis = getAnalysis(model, svc);
  const length = model.getValueLength();
  const markers: M.editor.IMarkerData[] = analysis.diagnostics.map((d) => {
    let start = d.start;
    let end = d.end;
    if (end <= start) {
      // Zero-width diagnostics (e.g. a missing operand at end of input) still need something to underline.
      if (start >= length && start > 0) start = length - 1;
      end = Math.min(length, start + 1);
    }
    const s = model.getPositionAt(start);
    const e = model.getPositionAt(end);
    return {
      severity: SEVERITY[d.severity],
      message: d.message,
      code: d.code,
      source: 'Cellsmith',
      startLineNumber: s.lineNumber,
      startColumn: s.column,
      endLineNumber: e.lineNumber,
      endColumn: e.column,
      tags: d.unnecessary ? [monaco.MarkerTag.Unnecessary] : undefined,
    };
  });
  monaco.editor.setModelMarkers(model, MARKER_OWNER, markers);
  return analysis;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const COMMON = new Set(
  (
    'IF IFS IFERROR SUM SUMIFS SUMIF COUNTIF COUNTIFS AVERAGE VLOOKUP XLOOKUP INDEX MATCH XMATCH LET LAMBDA FILTER SORT ' +
    'UNIQUE TEXT CONCAT TEXTJOIN LEFT RIGHT MID LEN TRIM AND OR NOT ROUND MAX MIN TODAY NOW DATE SEQUENCE MAP REDUCE ' +
    'TAKE DROP HSTACK VSTACK TEXTSPLIT ISBLANK ISNUMBER COUNTA SUBSTITUTE'
  ).split(' '),
);

function labelFor(name: string, params: { name: string; optional: boolean }[], repeat: boolean) {
  let label = name + '(';
  const offsets: [number, number][] = [];
  params.forEach((p, i) => {
    if (i > 0) label += ', ';
    const text = p.optional ? `[${p.name}]` : p.name;
    offsets.push([label.length, label.length + text.length]);
    label += text;
  });
  if (repeat) label += ', ...';
  return { label: label + ')', offsets };
}

function functionNotes(info: FunctionInfo): string {
  return [info.category, info.volatile ? 'volatile (recalculates every time the sheet does)' : ''].filter(Boolean).join(' · ');
}

/** Description only: signature help already shows the signature in its label. */
function functionDescription(info: FunctionInfo): M.IMarkdownString {
  return { value: `${info.description}\n\n_${functionNotes(info)}_` };
}

/** Full docs for completion details and hovers: signature, description, category. */
function functionDocs(info: FunctionInfo): M.IMarkdownString {
  return {
    value: `\`\`\`${LANGUAGE_ID}\n${info.signature}\n\`\`\`\n\n${info.description}\n\n_${functionNotes(info)}_`,
  };
}

function functionSnippet(info: FunctionInfo, hasParen: boolean): string {
  if (hasParen) return info.name;
  if (info.name === 'LET') return 'LET(${1:name1}, ${2:value1}, ${3:calculation_or_name2})';
  if (info.name === 'LAMBDA') return 'LAMBDA(${1:parameter1}, ${2:calculation})';
  if (info.maxArgs === 0) return `${info.name}()`;
  const required = info.params.filter((p) => !p.optional);
  if (required.length === 0) return `${info.name}($0)`;
  return `${info.name}(${required.map((p, i) => `\${${i + 1}:${p.name}}`).join(', ')})$0`;
}

function rangeOf(model: M.editor.ITextModel, start: number, end: number): M.IRange {
  const s = model.getPositionAt(start);
  const e = model.getPositionAt(end);
  return { startLineNumber: s.lineNumber, startColumn: s.column, endLineNumber: e.lineNumber, endColumn: e.column };
}

function trimSource(s: string, max = 300): string {
  const flat = s.replace(/\s+/g, ' ').trim();
  return flat.length > max ? flat.slice(0, max) + '…' : flat;
}

function symbolAt(analysis: Analysis, offset: number): { sym: Sym; span: { start: number; end: number }; isDecl: boolean } | undefined {
  for (const sym of analysis.symbols) {
    const d = sym.declaration;
    if (offset >= d.start && offset <= d.end) return { sym, span: d, isDecl: true };
    for (const u of sym.uses) if (offset >= u.start && offset <= u.end) return { sym, span: u, isDecl: false };
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

export function registerProviders(svc: EditorServices): M.IDisposable[] {
  const L = monaco.languages;
  const kinds = L.CompletionItemKind;
  const disposables: M.IDisposable[] = [];

  // --- Completion ---------------------------------------------------------
  disposables.push(
    L.registerCompletionItemProvider(LANGUAGE_ID, {
      triggerCharacters: ['[', '@', '#'],
      provideCompletionItems(model, position) {
        const text = model.getValue();
        const offset = model.getOffsetAt(position);
        const ctx = completionContext(text, offset);
        const wb = svc.getWorkbook();
        if (ctx.kind === 'none') return { suggestions: [] };

        if (ctx.kind === 'table') {
          const tables = wb?.tables ?? [];
          const table = ctx.table ? tables.find((t) => t.name.toLowerCase() === ctx.table!.toLowerCase()) : undefined;
          const columns = table ? table.columns : [...new Set(tables.flatMap((t) => t.columns))];
          const close = text[ctx.end] === ']' ? '' : ']';
          const range = rangeOf(model, ctx.start, ctx.end);
          const suggestions: M.languages.CompletionItem[] = columns.map((c) => ({
            label: c,
            kind: kinds.Field,
            detail: table ? `Column of ${table.name}` : 'Table column',
            insertText: escapeTableItem(c) + close,
            range,
          }));
          if (text[ctx.start - 1] !== '@') {
            for (const item of ['#All', '#Data', '#Headers', '#Totals', '#This Row']) {
              suggestions.push({ label: item, kind: kinds.Keyword, detail: 'Table area', insertText: item + close, range, sortText: 'z' + item });
            }
          }
          return { suggestions };
        }

        // Identifier completion
        const range = rangeOf(model, ctx.start, ctx.end);
        const hasParen = ctx.nextChar === '(';
        const suggestions: M.languages.CompletionItem[] = [];

        // LET variables and LAMBDA parameters in scope
        const analysis = getAnalysis(model, svc);
        for (const sym of symbolsAt(analysis, offset)) {
          const value = sym.value ? trimSource(text.slice(sym.value.start, sym.value.end), 80) : undefined;
          suggestions.push({
            label: sym.name,
            kind: sym.kind === 'parameter' ? kinds.TypeParameter : kinds.Variable,
            detail: sym.kind === 'parameter' ? 'LAMBDA parameter' : `LET variable${value ? ` = ${value}` : ''}`,
            insertText: sym.name,
            range,
            sortText: '0' + sym.name,
          });
        }

        if (wb) {
          const lambdas = new Map(namedLambdas(wb).map((l) => [l.name.toLowerCase(), l]));
          for (const n of wb.names) {
            const lam = lambdas.get(n.name.toLowerCase());
            if (lam) {
              const required = lam.params.filter((p) => !p.optional);
              suggestions.push({
                label: n.name,
                kind: kinds.Function,
                detail: labelFor(n.name, lam.params, false).label,
                documentation: { value: `Named LAMBDA\n\n\`\`\`${LANGUAGE_ID}\n${trimSource(n.refersTo, 400)}\n\`\`\`` },
                insertText: hasParen ? n.name : `${n.name}(${required.map((p, i) => `\${${i + 1}:${p.name}}`).join(', ')})$0`,
                insertTextRules: L.CompletionItemInsertTextRule.InsertAsSnippet,
                range,
                sortText: '1' + n.name,
              });
            } else {
              suggestions.push({
                label: n.name,
                kind: kinds.Constant,
                detail: `Name · ${trimSource(n.refersTo, 60)}`,
                insertText: n.name,
                range,
                sortText: '1' + n.name,
              });
            }
          }
          for (const t of wb.tables) {
            suggestions.push({
              label: t.name,
              kind: kinds.Struct,
              detail: `Table · ${t.columns.length} column${t.columns.length === 1 ? '' : 's'}`,
              documentation: { value: t.columns.map((c) => `- ${c}`).join('\n') },
              insertText: `${t.name}[$0]`,
              insertTextRules: L.CompletionItemInsertTextRule.InsertAsSnippet,
              command: { id: 'editor.action.triggerSuggest', title: 'Suggest columns' },
              range,
              sortText: '1' + t.name,
            });
          }
          for (const s of wb.sheets) {
            const q = quoteSheetName(s);
            suggestions.push({ label: s, kind: kinds.Module, detail: 'Sheet', insertText: `${q}!`, filterText: s, range, sortText: '4' + s });
          }
        }

        for (const info of allFunctions()) {
          suggestions.push({
            label: info.name,
            kind: kinds.Function,
            detail: info.signature,
            documentation: functionDocs(info),
            insertText: functionSnippet(info, hasParen),
            insertTextRules: L.CompletionItemInsertTextRule.InsertAsSnippet,
            // Pop up the argument hints as soon as a function is inserted, like VS Code does.
            command: hasParen || info.maxArgs === 0 ? undefined : { id: 'editor.action.triggerParameterHints', title: 'Signature help' },
            range,
            sortText: (COMMON.has(info.name) ? '2' : '3') + info.name,
          });
        }
        for (const c of ['TRUE', 'FALSE']) {
          suggestions.push({ label: c, kind: kinds.Keyword, insertText: c, range, sortText: '5' + c });
        }
        return { suggestions };
      },
    }),
  );

  // --- Signature help -----------------------------------------------------
  disposables.push(
    L.registerSignatureHelpProvider(LANGUAGE_ID, {
      signatureHelpTriggerCharacters: ['(', ','],
      signatureHelpRetriggerCharacters: [' '],
      provideSignatureHelp(model, position) {
        const analysis = getAnalysis(model, svc);
        const offset = model.getOffsetAt(position);
        const call = enclosingCall(analysis.parse, offset);
        if (!call) return null;
        const { fn, argIndex, argc } = call;

        let params: { name: string; optional: boolean }[] | undefined;
        let repeat = false;
        let doc: M.IMarkdownString | undefined;
        let active = argIndex;
        let nameForLabel = fn.name;

        const info = getFunction(fn.name);
        if (info) {
          params = info.params;
          repeat = info.repeat > 0 && info.name !== 'LAMBDA';
          doc = functionDescription(info);
          active = activeParamIndex(info, argIndex, argc);
          nameForLabel = info.name;
        } else {
          const local = symbolsAt(analysis, offset).find((s) => s.name.toLowerCase() === fn.name.toLowerCase());
          const lam = local?.value;
          if (lam?.type === 'Function' && lam.name.toUpperCase() === 'LAMBDA') {
            params = lam.args.slice(0, -1).map((a) =>
              a.type === 'Reference' && a.kind === 'table' ? { name: a.text.slice(1, -1), optional: true } : { name: a.type === 'Reference' ? a.text : '?', optional: false },
            );
            doc = { value: `LAMBDA defined by LET variable \`${local!.name}\`` };
          } else {
            const named = svc.getWorkbook() ? namedLambdas(svc.getWorkbook()!).find((l) => l.name.toLowerCase() === fn.name.toLowerCase()) : undefined;
            if (named) {
              params = named.params;
              doc = { value: `Named LAMBDA\n\n\`\`\`${LANGUAGE_ID}\n${trimSource(named.refersTo, 400)}\n\`\`\`` };
            }
          }
          if (params) active = Math.min(argIndex, Math.max(0, params.length - 1));
        }
        if (!params) return null;

        const { label, offsets } = labelFor(nameForLabel, params, repeat);
        return {
          value: {
            signatures: [{ label, documentation: doc, parameters: offsets.map((o) => ({ label: o })), activeParameter: active }],
            activeSignature: 0,
            activeParameter: active,
          },
          dispose() {},
        };
      },
    }),
  );

  // --- Hover --------------------------------------------------------------
  disposables.push(
    L.registerHoverProvider(LANGUAGE_ID, {
      provideHover(model, position) {
        const analysis = getAnalysis(model, svc);
        if (!analysis.parse.formula.hasEquals) return null;
        const text = model.getValue();
        const offset = model.getOffsetAt(position);

        const hit = symbolAt(analysis, offset);
        if (hit) {
          const { sym, span } = hit;
          const kind = sym.kind === 'parameter' ? 'LAMBDA parameter' : 'LET variable';
          const lines = [`**${sym.name}** · ${kind}${sym.optional ? ' (optional)' : ''}`];
          if (sym.value) lines.push(`\`\`\`${LANGUAGE_ID}\n= ${trimSource(text.slice(sym.value.start, sym.value.end))}\n\`\`\``);
          return { range: rangeOf(model, span.start, span.end), contents: [{ value: lines.join('\n\n') }] };
        }

        const path = nodePathAt(analysis.parse.formula.body, offset);
        const node = path[path.length - 1];
        if (!node) return null;
        if (node.type === 'Function' && offset >= node.nameSpan.start && offset <= node.nameSpan.end) {
          const info = getFunction(node.name);
          if (info) return { range: rangeOf(model, node.nameSpan.start, node.nameSpan.end), contents: [functionDocs(info)] };
          const named = namedLambdas(svc.getWorkbook() ?? { sheets: [], names: [], tables: [] }).find((l) => l.name.toLowerCase() === node.name.toLowerCase());
          if (named) {
            return {
              range: rangeOf(model, node.nameSpan.start, node.nameSpan.end),
              contents: [{ value: `**${named.name}** · named LAMBDA\n\n\`\`\`${LANGUAGE_ID}\n${trimSource(named.refersTo, 400)}\n\`\`\`` }],
            };
          }
        }
        if (node.type === 'Reference' && node.kind === 'name' && !node.prefix) {
          const def = svc.getWorkbook()?.names.find((n) => n.name.toLowerCase() === node.text.toLowerCase());
          if (def) {
            return {
              range: rangeOf(model, node.start, node.end),
              contents: [{ value: `**${def.name}** · defined name\n\n\`\`\`${LANGUAGE_ID}\n${trimSource(def.refersTo, 300)}\n\`\`\`` }],
            };
          }
        }
        return null;
      },
    }),
  );

  // --- Rename & go to definition (LET variables / LAMBDA parameters) -------
  const NAME_OK = /^[A-Za-z_\\][\w.\\]*$/;
  const CELL_LIKE = /^[A-Za-z]{1,3}\d+$/;
  disposables.push(
    L.registerRenameProvider(LANGUAGE_ID, {
      resolveRenameLocation(model, position) {
        const hit = symbolAt(getAnalysis(model, svc), model.getOffsetAt(position));
        if (!hit) return { range: new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column), text: '', rejectReason: 'Only LET variables and LAMBDA parameters can be renamed.' };
        const text = model.getValue().slice(hit.span.start, hit.span.end);
        const inner = text.startsWith('[') ? text.slice(1, -1) : text;
        const offset = text.startsWith('[') ? 1 : 0;
        return { range: rangeOf(model, hit.span.start + offset, hit.span.start + offset + inner.length), text: inner };
      },
      provideRenameEdits(model, position, newName) {
        const hit = symbolAt(getAnalysis(model, svc), model.getOffsetAt(position));
        if (!hit) return { edits: [], rejectReason: 'Nothing to rename here.' };
        if (!NAME_OK.test(newName) || CELL_LIKE.test(newName) || /^(true|false)$/i.test(newName)) {
          return { edits: [], rejectReason: `'${newName}' is not a valid name: use letters, digits, _ and . and do not look like a cell reference.` };
        }
        const text = model.getValue();
        const spans = [hit.sym.declaration, ...hit.sym.uses];
        return {
          edits: spans.map((s) => {
            const raw = text.slice(s.start, s.end);
            return {
              resource: model.uri,
              versionId: undefined,
              textEdit: { range: rangeOf(model, s.start, s.end), text: raw.startsWith('[') ? `[${newName}]` : newName },
            };
          }),
        };
      },
    }),
    L.registerDefinitionProvider(LANGUAGE_ID, {
      provideDefinition(model, position) {
        const hit = symbolAt(getAnalysis(model, svc), model.getOffsetAt(position));
        if (!hit || hit.isDecl) return null;
        return { uri: model.uri, range: rangeOf(model, hit.sym.declaration.start, hit.sym.declaration.end) };
      },
    }),
  );

  // --- Quick fixes --------------------------------------------------------
  disposables.push(
    L.registerCodeActionProvider(LANGUAGE_ID, {
      provideCodeActions(model, _range, context) {
        const analysis = getAnalysis(model, svc);
        const actions: M.languages.CodeAction[] = [];
        for (const marker of context.markers) {
          const start = model.getOffsetAt({ lineNumber: marker.startLineNumber, column: marker.startColumn });
          const diag = analysis.diagnostics.find((d) => d.code === marker.code && d.start === start && d.fix);
          if (!diag?.fix) continue;
          actions.push({
            title: diag.fix.title,
            kind: 'quickfix',
            diagnostics: [marker],
            isPreferred: true,
            edit: {
              edits: [
                {
                  resource: model.uri,
                  versionId: undefined,
                  textEdit: { range: rangeOf(model, diag.start, diag.end), text: diag.fix.text },
                },
              ],
            },
          });
        }
        return { actions, dispose() {} };
      },
    }),
  );

  // --- Formatting (Shift+Alt+F) --------------------------------------------
  disposables.push(
    L.registerDocumentFormattingEditProvider(LANGUAGE_ID, {
      provideDocumentFormattingEdits(model) {
        const result = formatFormula(model.getValue(), svc.getFormatOptions());
        if (!result.ok || result.text === model.getValue()) return [];
        return [{ range: model.getFullModelRange(), text: result.text }];
      },
    }),
  );

  return disposables;
}
