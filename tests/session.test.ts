import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatFormula, minifyFormula } from '../src/core/formatter';
import { analyze } from '../src/core/analysis';
import type { WorkbookInfo } from '../src/editor/context';
import { Session, type SessionState } from '../src/excel/session';
import { BusyError, type CellTarget, type SelectionInfo, type WorkbookAdapter } from '../src/excel/types';
import { DEFAULT_SETTINGS, type Settings } from '../src/settings';

class MockAdapter implements WorkbookAdapter {
  readonly kind = 'excel' as const;
  cells = new Map<string, string>();
  selected: CellTarget = { sheet: 'Sheet1', address: 'A1' };
  handler?: () => void;
  writes: { target: CellTarget; text: string }[] = [];
  busyWrites = 0;
  rejectNext?: string;

  set(addr: string, text: string) {
    this.cells.set(addr, text);
  }
  select(addr: string) {
    this.selected = { sheet: 'Sheet1', address: addr };
    this.handler?.();
  }
  async getSelection(): Promise<SelectionInfo> {
    const text = this.cells.get(this.selected.address) ?? '';
    return { target: this.selected, text, display: `shown(${text})`, cellCount: 1, isFormula: text.startsWith('=') };
  }
  async onSelectionChanged(h: () => void) {
    this.handler = h;
  }
  async write(target: CellTarget, text: string) {
    if (this.busyWrites > 0) {
      this.busyWrites--;
      throw new BusyError();
    }
    if (this.rejectNext) {
      const m = this.rejectNext;
      this.rejectNext = undefined;
      throw new Error(m);
    }
    this.writes.push({ target, text });
    this.cells.set(target.address, text);
    return `shown(${text})`;
  }
  async getWorkbookInfo(): Promise<WorkbookInfo> {
    return { sheets: ['Sheet1'], names: [], tables: [] };
  }
}

function setup(overrides: Partial<Settings> = {}) {
  const adapter = new MockAdapter();
  let text = '';
  let settings: Settings = { ...DEFAULT_SETTINGS, formatOnLoad: false, ...overrides };
  const states: SessionState[] = [];
  const session = new Session({
    adapter,
    settings: () => settings,
    getText: () => text,
    setText: (t) => {
      text = t;
    },
    countErrors: (t) => analyze(t).diagnostics.filter((d) => d.severity === 'error').length,
    format: (t) => {
      const r = formatFormula(t);
      return r.ok ? r.text : undefined;
    },
    minify: (t) => {
      const r = minifyFormula(t);
      return r.ok ? r.text : undefined;
    },
    onState: (s) => states.push(s),
    onWorkbook: () => {},
  });
  /** Simulates the user typing into the editor. */
  const type = (t: string) => {
    text = t;
    session.handleEdit();
  };
  return {
    adapter,
    session,
    states,
    type,
    editor: () => text,
    last: () => states[states.length - 1]!,
    setSettings: (s: Partial<Settings>) => (settings = { ...settings, ...s }),
  };
}

beforeEach(() => vi.useFakeTimers());
afterEach(() => vi.useRealTimers());

describe('Session', () => {
  it('loads the selected cell into the editor without marking it dirty', async () => {
    const t = setup();
    t.adapter.set('A1', '=SUM(B1:B3)');
    await t.session.start();
    expect(t.editor()).toBe('=SUM(B1:B3)');
    expect(t.session.isDirty()).toBe(false);
    expect(t.last()).toMatchObject({ phase: 'synced', target: { address: 'A1' }, display: 'shown(=SUM(B1:B3))' });
    expect(t.adapter.writes).toEqual([]);
  });

  it('pretty-prints single-line formulas on load but never writes them back by itself', async () => {
    const t = setup({ formatOnLoad: true });
    t.adapter.set('A1', '=LET(x,1,y,2,x+y)');
    await t.session.start();
    expect(t.editor()).toBe('=LET(\n    x, 1,\n    y, 2,\n    x + y\n)');
    expect(t.session.isDirty()).toBe(false);
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toEqual([]);
  });

  it('leaves formulas the user already laid out across lines alone', async () => {
    const t = setup({ formatOnLoad: true });
    t.adapter.set('A1', '=IF(\n  A1,\n  1,\n  2\n)');
    await t.session.start();
    expect(t.editor()).toBe('=IF(\n  A1,\n  1,\n  2\n)');
  });

  it('auto-applies a valid edit after a pause, once, as typed', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=SUM(A1:A3)');
    expect(t.last().phase).toBe('dirty');
    await vi.advanceTimersByTimeAsync(300);
    expect(t.adapter.writes).toEqual([]); // still within the debounce window
    await vi.advanceTimersByTimeAsync(600);
    expect(t.adapter.writes).toEqual([{ target: { sheet: 'Sheet1', address: 'A1' }, text: '=SUM(A1:A3)' }]);
    expect(t.last().phase).toBe('applied');
    expect(t.session.isDirty()).toBe(false);
  });

  it('debounces: only the last of several quick edits is written', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    for (const s of ['=S', '=SU', '=SUM(', '=SUM(1)']) {
      t.type(s);
      await vi.advanceTimersByTimeAsync(100);
    }
    await vi.advanceTimersByTimeAsync(1000);
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=SUM(1)']);
  });

  it('does not write formulas that have errors', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=SUM(1,');
    expect(t.last().phase).toBe('invalid');
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toEqual([]);
  });

  it('does not write arity errors either', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=IF()');
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toEqual([]);
    expect(t.last().phase).toBe('invalid');
  });

  it('waits for an explicit apply when auto-apply is off', async () => {
    const t = setup({ autoApply: false });
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=2');
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toEqual([]);
    await t.session.applyNow();
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=2']);
  });

  it('refuses an explicit apply while there are errors', async () => {
    const t = setup({ autoApply: false });
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=(');
    await t.session.applyNow();
    expect(t.adapter.writes).toEqual([]);
    expect(t.last().phase).toBe('invalid');
  });

  it('flushes a pending valid edit to the PREVIOUS cell when the selection changes', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    t.adapter.set('B1', '=2');
    await t.session.start();
    t.type('=SUM(5)');
    t.adapter.select('B1');
    await vi.advanceTimersByTimeAsync(10);
    expect(t.adapter.writes).toEqual([{ target: { sheet: 'Sheet1', address: 'A1' }, text: '=SUM(5)' }]);
    expect(t.editor()).toBe('=2');
    expect(t.last().target?.address).toBe('B1');
    // the timer from the first edit must not fire against the new cell
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toHaveLength(1);
  });

  it('discards an invalid pending edit on selection change and says so', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    t.adapter.set('B1', '=2');
    await t.session.start();
    t.type('=SUM(');
    const seen: string[] = [];
    const orig = t.session as unknown as { emit: (p: { message?: string }) => void };
    const emit = orig.emit.bind(t.session);
    orig.emit = (p) => {
      if (p.message) seen.push(p.message);
      emit(p);
    };
    t.adapter.select('B1');
    await vi.advanceTimersByTimeAsync(10);
    expect(t.adapter.writes).toEqual([]);
    expect(seen.some((m) => /Discarded/.test(m))).toBe(true);
    expect(t.editor()).toBe('=2');
  });

  it('keeps unsaved typing when the same cell is re-selected', async () => {
    const t = setup({ autoApply: false });
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=1+1');
    t.adapter.select('A1');
    await vi.advanceTimersByTimeAsync(10);
    expect(t.editor()).toBe('=1+1');
  });

  it('retries when Excel is busy in cell-edit mode', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.adapter.busyWrites = 1;
    t.type('=2');
    await vi.advanceTimersByTimeAsync(800);
    expect(t.last().phase).toBe('busy');
    await vi.advanceTimersByTimeAsync(2000);
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=2']);
    expect(t.last().phase).toBe('applied');
  });

  it('surfaces Excel rejections without losing the edit', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.adapter.rejectNext = 'Excel rejected the formula.';
    t.type('=2');
    await vi.advanceTimersByTimeAsync(1000);
    expect(t.last()).toMatchObject({ phase: 'error', message: 'Excel rejected the formula.' });
    expect(t.editor()).toBe('=2');
    expect(t.session.isDirty()).toBe(true);
  });

  it('reverting to the original text is not an edit', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=2');
    t.type('=1');
    expect(t.last().phase).toBe('synced');
    await vi.advanceTimersByTimeAsync(5000);
    expect(t.adapter.writes).toEqual([]);
  });

  it('minified write mode strips whitespace from the cell but not the editor', async () => {
    const t = setup({ writeMode: 'minified' });
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=IF(A1 > 0,\n    "yes",\n    "no")');
    await vi.advanceTimersByTimeAsync(1000);
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=IF(A1>0,"yes","no")']);
    expect(t.editor()).toContain('\n');
  });

  it('normalises line endings and trailing whitespace on write', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=IF(\r\n  A1,\r\n  1,\r\n  2\r\n)\r\n\r\n');
    await vi.advanceTimersByTimeAsync(1000);
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=IF(\n  A1,\n  1,\n  2\n)']);
  });

  it('reload discards edits instead of writing them', async () => {
    const t = setup({ autoApply: false });
    t.adapter.set('A1', '=1');
    await t.session.start();
    t.type('=999');
    await t.session.reload();
    expect(t.editor()).toBe('=1');
    expect(t.adapter.writes).toEqual([]);
  });

  it('applies edits typed while a write is in flight', async () => {
    const t = setup();
    t.adapter.set('A1', '=1');
    await t.session.start();
    const realWrite = t.adapter.write.bind(t.adapter);
    let first = true;
    t.adapter.write = async (target, text) => {
      if (first) {
        first = false;
        t.type('=3'); // user keeps typing during the write
      }
      return realWrite(target, text);
    };
    t.type('=2');
    await vi.advanceTimersByTimeAsync(800);
    await vi.advanceTimersByTimeAsync(1000);
    expect(t.adapter.writes.map((w) => w.text)).toEqual(['=2', '=3']);
    expect(t.session.isDirty()).toBe(false);
  });
});
