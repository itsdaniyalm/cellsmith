import './style.css';
import type * as M from 'monaco-editor';
import { analyze } from '../core/analysis';
import { DEFAULT_FORMAT_OPTIONS, formatFormula, minifyFormula, type FormatOptions } from '../core/formatter';
import { lintContextFor, type WorkbookInfo } from '../editor/context';
import { LANGUAGE_ID, THEME_DARK, THEME_LIGHT, registerFormulaLanguage } from '../editor/language';
import { monaco } from '../editor/monaco';
import { getAnalysis, registerProviders, validateModel, type EditorServices } from '../editor/providers';
import { OfficeAdapter } from '../excel/office';
import { Session, type Phase, type SessionState } from '../excel/session';
import { StandaloneAdapter } from '../excel/standalone';
import { formatTarget, type WorkbookAdapter } from '../excel/types';
import { loadSettings, saveSettings, type Settings } from '../settings';

const $ = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing #${id}`);
  return el as T;
};

// ---------------------------------------------------------------------------
// Environment
// ---------------------------------------------------------------------------

/** True when running inside Excel; false in an ordinary browser tab (preview mode). */
async function insideExcel(): Promise<boolean> {
  if (typeof Office === 'undefined') return false;
  const ready = Office.onReady().then((info) => info.host === Office.HostType.Excel);
  const timeout = new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 4000));
  return Promise.race([ready, timeout]);
}

function prefersDark(): boolean {
  try {
    const bg = Office.context?.officeTheme?.bodyBackgroundColor;
    const m = bg ? /^#?([0-9a-f]{6})$/i.exec(bg) : null;
    if (m) {
      const n = parseInt(m[1]!, 16);
      const luminance = (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
      return luminance < 0.5;
    }
  } catch {
    /* no Office theme available */
  }
  return window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false;
}

function applyTheme(): void {
  const dark = prefersDark();
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  monaco.editor.setTheme(dark ? THEME_DARK : THEME_LIGHT);
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

async function boot() {
  const connected = await insideExcel();
  const adapter: WorkbookAdapter = connected ? new OfficeAdapter() : new StandaloneAdapter();
  let settings: Settings = loadSettings();

  registerFormulaLanguage();
  applyTheme();
  window.matchMedia?.('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);

  // Workbook metadata (sheets, names, tables) feeds completion and linting. Until the first read
  // completes it stays undefined, which switches off "undefined name" checks rather than flagging everything.
  let workbook: WorkbookInfo | undefined;
  let workbookVersion = 0;

  const model = monaco.editor.createModel('', LANGUAGE_ID, monaco.Uri.parse('inmemory://cellsmith/formula'));
  model.setEOL(monaco.editor.EndOfLineSequence.LF);

  const editor = monaco.editor.create($('editor'), {
    model,
    theme: prefersDark() ? THEME_DARK : THEME_LIGHT,
    fontFamily: "Consolas, 'Cascadia Mono', 'Courier New', monospace",
    fontSize: 13,
    lineHeight: 20,
    minimap: { enabled: false },
    lineNumbers: 'on',
    lineNumbersMinChars: 2,
    glyphMargin: false,
    folding: true,
    wordWrap: 'on',
    wrappingIndent: 'indent',
    scrollBeyondLastLine: false,
    automaticLayout: true,
    tabSize: 4,
    insertSpaces: true,
    detectIndentation: false,
    bracketPairColorization: { enabled: true },
    guides: { bracketPairs: true, indentation: true },
    matchBrackets: 'always',
    quickSuggestions: { other: true, comments: false, strings: false },
    suggestOnTriggerCharacters: true,
    snippetSuggestions: 'top',
    tabCompletion: 'on',
    parameterHints: { enabled: true, cycle: true },
    fixedOverflowWidgets: true,
    renderLineHighlight: 'line',
    padding: { top: 8, bottom: 8 },
    overviewRulerLanes: 0,
    hideCursorInOverviewRuler: true,
    scrollbar: { horizontal: 'hidden', verticalScrollbarSize: 10 },
    placeholder: connected ? 'Select a cell in Excel to edit its formula' : 'Type a formula, starting with =',
    ariaLabel: 'Excel formula editor',
  });

  /** Width to format to: a fixed number, or as many columns as fit in the pane. */
  const formatOptions = (): Partial<FormatOptions> => {
    let width: number;
    if (settings.width === 'auto') {
      const layout = editor.getLayoutInfo();
      const charWidth = editor.getOption(monaco.editor.EditorOption.fontInfo).typicalHalfwidthCharacterWidth;
      // Before the first layout the pane reports zero width; fall back to a sensible default.
      const usable = layout.contentWidth - layout.verticalScrollbarWidth;
      width = usable > 150 && charWidth > 0 ? Math.floor(usable / charWidth) - 2 : 80;
      width = Math.max(30, Math.min(100, width));
    } else {
      width = settings.width;
    }
    return {
      width,
      indent: settings.indent,
      uppercaseFunctions: settings.uppercaseFunctions,
      expandLetAt: settings.expandLetAt,
    };
  };

  const services: EditorServices = {
    getWorkbook: () => workbook,
    workbookVersion: () => workbookVersion,
    getFormatOptions: formatOptions,
  };
  registerProviders(services);

  // -------------------------------------------------------------------------
  // Problems list
  // -------------------------------------------------------------------------

  const problemsEl = $('problems');
  function renderProblems(): void {
    const analysis = getAnalysis(model, services);
    const items = analysis.diagnostics.filter((d) => d.severity !== 'hint');
    problemsEl.replaceChildren();
    if (items.length === 0) {
      if (analysis.parse.formula.hasEquals) {
        const ok = document.createElement('div');
        ok.className = 'no-problems';
        ok.textContent = 'No problems found.';
        problemsEl.append(ok);
      }
      return;
    }
    for (const d of items) {
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'problem';
      row.dataset.sev = d.severity;
      const pos = model.getPositionAt(d.start);

      const dot = document.createElement('span');
      dot.className = 'dot';
      const msg = document.createElement('span');
      msg.className = 'msg';
      msg.textContent = d.message;
      const where = document.createElement('span');
      where.className = 'pos';
      where.textContent = `${pos.lineNumber}:${pos.column}`;
      row.append(dot, msg, where);

      row.addEventListener('click', () => {
        const end = model.getPositionAt(Math.max(d.end, d.start));
        editor.setSelection(new monaco.Selection(pos.lineNumber, pos.column, end.lineNumber, end.column));
        editor.revealPositionInCenterIfOutsideViewport(pos);
        editor.focus();
      });
      if (d.fix) {
        const fix = document.createElement('span');
        fix.className = 'fix';
        fix.textContent = d.fix.text;
        fix.title = d.fix.title;
        fix.setAttribute('role', 'button');
        fix.addEventListener('click', (ev) => {
          ev.stopPropagation();
          const s = model.getPositionAt(d.start);
          const e = model.getPositionAt(d.end);
          editor.executeEdits('quick-fix', [{ range: new monaco.Range(s.lineNumber, s.column, e.lineNumber, e.column), text: d.fix!.text }]);
        });
        row.insertBefore(fix, where);
      }
      problemsEl.append(row);
    }
  }

  function revalidate(): void {
    validateModel(model, services);
    renderProblems();
  }

  // -------------------------------------------------------------------------
  // Session (sync with Excel)
  // -------------------------------------------------------------------------

  let suppressEdit = false;
  const BADGE: Record<Phase, string> = {
    idle: 'No cell',
    synced: 'In sync',
    dirty: 'Editing',
    invalid: 'Has errors',
    applying: 'Applying…',
    applied: 'Applied',
    busy: 'Excel busy',
    error: 'Not applied',
  };

  const chip = $('cell-chip');
  const badge = $('badge');
  const status = $('status');
  const resultBox = $('result');
  const resultValue = $('result-value');

  function renderState(s: SessionState): void {
    chip.textContent = s.target ? formatTarget(s.target) : 'No cell';
    badge.dataset.phase = s.phase;
    badge.textContent = BADGE[s.phase];
    status.textContent = s.message || (connected ? '' : 'Preview mode: not connected to Excel. Edits stay on this page.');
    status.classList.toggle('is-error', s.phase === 'error' || s.phase === 'invalid');
    const showResult = connected && s.display !== '' && s.phase !== 'idle';
    resultBox.hidden = !showResult;
    resultValue.textContent = s.display;
    resultValue.title = s.display;
    if (s.phase === 'synced') void session.refreshWorkbook();
  }

  const session = new Session({
    adapter,
    settings: () => settings,
    getText: () => model.getValue(),
    setText: (text) => {
      suppressEdit = true;
      try {
        model.setValue(text);
      } finally {
        suppressEdit = false;
      }
      revalidate();
    },
    countErrors: (text) => analyze(text, lintContextFor(workbook)).diagnostics.filter((d) => d.severity === 'error').length,
    format: (text) => {
      const r = formatFormula(text, formatOptions());
      return r.ok ? r.text : undefined;
    },
    minify: (text) => {
      const r = minifyFormula(text, formatOptions());
      return r.ok ? r.text : undefined;
    },
    onState: renderState,
    onWorkbook: (info) => {
      workbook = info; // preview mode receives a demo workbook so completion has something to show
      workbookVersion++;
      revalidate();
    },
  });

  model.onDidChangeContent(() => {
    if (suppressEdit) return;
    revalidate();
    session.handleEdit();
  });

  // -------------------------------------------------------------------------
  // Commands & controls
  // -------------------------------------------------------------------------

  const replaceAll = (text: string) => {
    editor.pushUndoStop();
    editor.executeEdits('cellsmith', [{ range: model.getFullModelRange(), text }]);
    editor.pushUndoStop();
  };

  const runFormat = () => {
    const r = formatFormula(model.getValue(), formatOptions());
    if (!r.ok) {
      status.textContent = 'Fix the syntax errors before formatting.';
      status.classList.add('is-error');
      return;
    }
    if (r.text !== model.getValue()) replaceAll(r.text);
  };
  const runMinify = () => {
    const r = minifyFormula(model.getValue(), formatOptions());
    if (!r.ok) {
      status.textContent = 'Fix the syntax errors before minifying.';
      status.classList.add('is-error');
      return;
    }
    if (r.text !== model.getValue()) replaceAll(r.text);
  };

  editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter, () => void session.applyNow());
  editor.addAction({ id: 'cellsmith.minify', label: 'Cellsmith: Minify Formula', run: runMinify });
  editor.addAction({ id: 'cellsmith.apply', label: 'Cellsmith: Apply to Cell', run: () => void session.applyNow() });

  $('btn-format').addEventListener('click', runFormat);
  $('btn-minify').addEventListener('click', runMinify);
  $('btn-apply').addEventListener('click', () => void session.applyNow());
  $('btn-reload').addEventListener('click', () => void session.reload());

  const settingsPanel = $('settings');
  const settingsBtn = $('btn-settings');
  settingsBtn.addEventListener('click', () => {
    settingsPanel.hidden = !settingsPanel.hidden;
    settingsBtn.setAttribute('aria-expanded', String(!settingsPanel.hidden));
  });

  const autoBox = $<HTMLInputElement>('chk-auto');
  const widthSel = $<HTMLSelectElement>('set-width');
  const indentSel = $<HTMLSelectElement>('set-indent');
  const letSel = $<HTMLSelectElement>('set-let');
  const writeSel = $<HTMLSelectElement>('set-write');
  const upperBox = $<HTMLInputElement>('set-upper');
  const formatLoadBox = $<HTMLInputElement>('set-formatload');

  const syncControls = () => {
    autoBox.checked = settings.autoApply;
    widthSel.value = String(settings.width);
    indentSel.value = String(settings.indent);
    letSel.value = String(settings.expandLetAt >= 999 ? 999 : settings.expandLetAt);
    writeSel.value = settings.writeMode;
    upperBox.checked = settings.uppercaseFunctions;
    formatLoadBox.checked = settings.formatOnLoad;
  };
  const update = (patch: Partial<Settings>) => {
    settings = { ...settings, ...patch };
    saveSettings(settings);
    syncControls();
  };
  syncControls();
  autoBox.addEventListener('change', () => {
    update({ autoApply: autoBox.checked });
    session.handleEdit();
  });
  widthSel.addEventListener('change', () => update({ width: widthSel.value === 'auto' ? 'auto' : Number(widthSel.value) }));
  indentSel.addEventListener('change', () => update({ indent: Number(indentSel.value) === 2 ? 2 : 4 }));
  letSel.addEventListener('change', () => update({ expandLetAt: Number(letSel.value) }));
  writeSel.addEventListener('change', () => update({ writeMode: writeSel.value === 'minified' ? 'minified' : 'as-typed' }));
  upperBox.addEventListener('change', () => update({ uppercaseFunctions: upperBox.checked }));
  formatLoadBox.addEventListener('change', () => update({ formatOnLoad: formatLoadBox.checked }));

  window.addEventListener('focus', () => void session.refreshWorkbook());

  // Handy for debugging in the browser console and for end-to-end checks.
  (window as unknown as { cellsmith: unknown }).cellsmith = { editor, model, session, monaco: monaco as typeof M };

  // Measure fonts and layout before the first cell loads, so "fit pane" formatting sees the real width.
  try {
    await document.fonts?.ready;
  } catch {
    /* font loading API unavailable */
  }
  monaco.editor.remeasureFonts();
  editor.layout();
  await session.start();
  revalidate();
}

boot().catch((err) => {
  console.error(err);
  const el = document.getElementById('status');
  if (el) {
    el.textContent = `Failed to start: ${err instanceof Error ? err.message : String(err)}`;
    el.classList.add('is-error');
  }
});
