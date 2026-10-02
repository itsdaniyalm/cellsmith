import type { WorkbookInfo } from '../editor/context';

export interface CellTarget {
  sheet: string;
  /** Single-cell A1 address without `$`, e.g. `B2`. */
  address: string;
}

export interface SelectionInfo {
  /** The top-left cell of the selection (the one we edit). */
  target: CellTarget;
  /** Formula text (`=SUM(A1:A3)`) or the constant's text. */
  text: string;
  /** What the cell displays, e.g. `42` or `#N/A`. */
  display: string;
  /** Number of cells in the selection. */
  cellCount: number;
  isFormula: boolean;
}

/** Thrown when Excel cannot service a request right now (typically: a cell is in edit mode). */
export class BusyError extends Error {
  constructor(message = 'Excel is busy: finish editing the cell and try again.') {
    super(message);
    this.name = 'BusyError';
  }
}

export interface WorkbookAdapter {
  readonly kind: 'excel' | 'standalone';
  getSelection(): Promise<SelectionInfo>;
  /** Subscribe to selection changes; resolves once the subscription is active. */
  onSelectionChanged(handler: () => void): Promise<void>;
  /** Write text (a formula or a constant) into a cell; resolves with the cell's new display text. */
  write(target: CellTarget, text: string): Promise<string>;
  getWorkbookInfo(): Promise<WorkbookInfo>;
}

export function formatTarget(t: CellTarget): string {
  const needsQuotes = !/^[A-Za-z_][\w.]*$/.test(t.sheet);
  return `${needsQuotes ? `'${t.sheet.replace(/'/g, "''")}'` : t.sheet}!${t.address}`;
}
