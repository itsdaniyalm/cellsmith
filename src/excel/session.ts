import type { WorkbookInfo } from '../editor/context';
import type { Settings } from '../settings';
import { BusyError, type CellTarget, type WorkbookAdapter } from './types';

export type Phase =
  | 'idle' // nothing loaded yet
  | 'synced' // editor matches the cell
  | 'dirty' // edited, waiting to be applied
  | 'invalid' // edited, but has errors so it cannot be applied
  | 'applying'
  | 'applied' // just written
  | 'busy' // Excel is in cell-edit mode; will retry
  | 'error'; // Excel rejected the write or a read failed

export interface SessionState {
  phase: Phase;
  target?: CellTarget;
  cellCount: number;
  /** Displayed value of the cell after the last read/write. */
  display: string;
  message: string;
}

export interface SessionDeps {
  adapter: WorkbookAdapter;
  settings(): Settings;
  /** Current editor text. */
  getText(): string;
  /** Replace the editor text programmatically (must not be reported back as a user edit). */
  setText(text: string): void;
  /** Number of error-severity problems in `text`. */
  countErrors(text: string): number;
  /** Pretty-print; returns undefined if the text cannot be formatted. */
  format(text: string): string | undefined;
  /** Strip whitespace; returns undefined if the text cannot be minified. */
  minify(text: string): string | undefined;
  onState(state: SessionState): void;
  onWorkbook(info: WorkbookInfo): void;
}

const AUTO_APPLY_DELAY_MS = 700;
const BUSY_RETRY_MS = 1500;
const WORKBOOK_REFRESH_MS = 10_000;

const sameTarget = (a?: CellTarget, b?: CellTarget) => !!a && !!b && a.sheet === b.sheet && a.address === b.address;

/**
 * Keeps the editor and the selected Excel cell in sync.
 *
 * - Selecting a cell loads its formula (pretty-printed if it is a single line).
 * - Edits are written back after a short pause, but only when they have no errors, since Excel
 *   rejects malformed formulas.
 * - Changing selection first flushes a pending valid edit to the *previous* cell, so nothing lands
 *   in the wrong place.
 */
export class Session {
  private state: SessionState = { phase: 'idle', cellCount: 0, display: '', message: '' };
  private baseline = '';
  private target?: CellTarget;
  private timer?: ReturnType<typeof setTimeout>;
  private chain: Promise<void> = Promise.resolve();
  private lastWorkbookRefresh = 0;

  constructor(private readonly deps: SessionDeps) {}

  async start(): Promise<void> {
    await this.deps.adapter.onSelectionChanged(() => void this.enqueue(() => this.load()));
    await this.enqueue(() => this.load());
    await this.refreshWorkbook(true);
  }

  /** Call when the user (not the app) changes the editor text. */
  handleEdit(): void {
    clearTimeout(this.timer);
    if (!this.target) return;
    if (!this.isDirty()) return this.emit({ phase: 'synced', message: '' });
    const errors = this.deps.countErrors(this.deps.getText());
    if (errors > 0) {
      return this.emit({ phase: 'invalid', message: `${errors} problem${errors === 1 ? '' : 's'} to fix before this can be applied.` });
    }
    if (this.deps.settings().autoApply) {
      this.emit({ phase: 'dirty', message: 'Applying…' });
      this.timer = setTimeout(() => void this.enqueue(() => this.applyIfDirty()), AUTO_APPLY_DELAY_MS);
    } else {
      this.emit({ phase: 'dirty', message: 'Unapplied changes. Press Ctrl+Enter to apply.' });
    }
  }

  /** Apply immediately (Ctrl+Enter / Apply button). */
  applyNow(): Promise<void> {
    clearTimeout(this.timer);
    return this.enqueue(async () => {
      if (!this.target) return;
      const errors = this.deps.countErrors(this.deps.getText());
      if (errors > 0) {
        return this.emit({ phase: 'invalid', message: `Fix ${errors} problem${errors === 1 ? '' : 's'} first: Excel would reject this formula.` });
      }
      await this.apply();
    });
  }

  /** Discard edits and reload the selected cell. */
  reload(): Promise<void> {
    clearTimeout(this.timer);
    return this.enqueue(() => this.load(true));
  }

  /** The text that would be written to the cell for the current editor content. */
  cellText(): string {
    const text = this.deps.getText().replace(/\r\n/g, '\n').replace(/\s+$/, '');
    if (this.deps.settings().writeMode === 'minified') return this.deps.minify(text) ?? text;
    return text;
  }

  isDirty(): boolean {
    return this.deps.getText() !== this.baseline;
  }

  async refreshWorkbook(force = false): Promise<void> {
    const now = Date.now();
    if (!force && now - this.lastWorkbookRefresh < WORKBOOK_REFRESH_MS) return;
    this.lastWorkbookRefresh = now;
    try {
      this.deps.onWorkbook(await this.deps.adapter.getWorkbookInfo());
    } catch {
      /* metadata is a nicety; completion just keeps the last known list */
    }
  }

  // -------------------------------------------------------------------------

  private enqueue(fn: () => Promise<void>): Promise<void> {
    this.chain = this.chain.then(fn, fn).catch(() => undefined);
    return this.chain;
  }

  private emit(patch: Partial<SessionState>): void {
    this.state = { ...this.state, target: this.target, ...patch };
    this.deps.onState(this.state);
  }

  /** Write a valid pending edit to its cell, or report why it was dropped. */
  private async flush(): Promise<void> {
    clearTimeout(this.timer);
    if (!this.target || !this.isDirty()) return;
    if (this.deps.countErrors(this.deps.getText()) === 0) {
      await this.apply();
    } else {
      this.emit({ phase: 'idle', message: `Discarded an unapplied edit to ${this.target.address}: it had errors.` });
    }
  }

  /** Timer callbacks run later, by which time the selection may have moved on and the text been replaced. */
  private async applyIfDirty(): Promise<void> {
    if (this.isDirty()) await this.apply();
  }

  private async apply(): Promise<void> {
    const target = this.target;
    if (!target) return;
    const snapshot = this.deps.getText();
    this.emit({ phase: 'applying', message: 'Writing to Excel…' });
    try {
      const display = await this.deps.adapter.write(target, this.cellText());
      this.baseline = snapshot;
      this.emit({ phase: 'applied', display, message: 'Applied.' });
      // The user kept typing while the write was in flight.
      if (this.isDirty()) this.handleEdit();
    } catch (err) {
      if (err instanceof BusyError) {
        this.emit({ phase: 'busy', message: err.message });
        this.timer = setTimeout(() => void this.enqueue(() => this.applyIfDirty()), BUSY_RETRY_MS);
      } else {
        this.emit({ phase: 'error', message: err instanceof Error ? err.message : String(err) });
      }
    }
  }

  private async load(force = false): Promise<void> {
    if (!force) await this.flush();
    let sel;
    try {
      sel = await this.deps.adapter.getSelection();
    } catch (err) {
      this.emit({ phase: err instanceof BusyError ? 'busy' : 'error', message: err instanceof Error ? err.message : String(err) });
      return;
    }
    // Re-selecting the cell we are editing must not clobber unsaved typing.
    if (!force && sameTarget(this.target, sel.target) && this.isDirty()) return;

    this.target = sel.target;
    let shown = sel.text;
    if (this.deps.settings().formatOnLoad && sel.isFormula && !sel.text.includes('\n')) {
      shown = this.deps.format(sel.text) ?? sel.text;
    }
    this.baseline = shown;
    if (this.deps.getText() !== shown) this.deps.setText(shown);
    this.emit({
      phase: 'synced',
      cellCount: sel.cellCount,
      display: sel.display,
      message: sel.cellCount > 1 ? `Editing the first of ${sel.cellCount} selected cells.` : '',
    });
  }
}
