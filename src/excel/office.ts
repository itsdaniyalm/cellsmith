import type { WorkbookInfo } from '../editor/context';
import { BusyError, type CellTarget, type SelectionInfo, type WorkbookAdapter } from './types';

/** `Sheet1!B2`, `'My Sheet'!$B$2` -> { sheet, address }. */
export function splitAddress(full: string): CellTarget {
  const bang = full.lastIndexOf('!');
  let sheet = bang < 0 ? '' : full.slice(0, bang);
  const address = (bang < 0 ? full : full.slice(bang + 1)).replace(/\$/g, '');
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, address };
}

/** Normalise Office.js failures into something the UI can show. */
function translate(err: unknown): Error {
  const e = err as { code?: string; message?: string };
  if (e?.code === 'InvalidOperationInCellEditMode' || /edit mode/i.test(e?.message ?? '')) return new BusyError();
  if (e?.code === 'InvalidArgument' || e?.code === 'GeneralException') {
    return new Error(`Excel rejected the formula${e.message ? `: ${e.message}` : '.'}`);
  }
  return err instanceof Error ? err : new Error(String(err));
}

const toText = (v: unknown): string => (v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));

export class OfficeAdapter implements WorkbookAdapter {
  readonly kind = 'excel' as const;

  async getSelection(): Promise<SelectionInfo> {
    try {
      return await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load(['rowCount', 'columnCount']);
        const cell = selection.getCell(0, 0);
        cell.load(['address', 'formulas', 'text']);
        await context.sync();
        const text = toText(cell.formulas[0]?.[0]);
        return {
          target: splitAddress(cell.address),
          text,
          display: toText(cell.text[0]?.[0]),
          cellCount: selection.rowCount * selection.columnCount,
          isFormula: text.startsWith('='),
        };
      });
    } catch (err) {
      throw translate(err);
    }
  }

  async onSelectionChanged(handler: () => void): Promise<void> {
    await Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(async () => handler());
      await context.sync();
    });
  }

  async write(target: CellTarget, text: string): Promise<string> {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getItem(target.sheet).getRange(target.address);
        range.formulas = [[text]];
        range.load('text');
        await context.sync();
        return toText(range.text[0]?.[0]);
      });
    } catch (err) {
      throw translate(err);
    }
  }

  async getWorkbookInfo(): Promise<WorkbookInfo> {
    try {
      return await Excel.run(async (context) => {
        const { workbook } = context;
        workbook.worksheets.load('items/name');
        workbook.names.load('items/name,items/formula');
        workbook.tables.load('items/name');
        await context.sync();

        for (const ws of workbook.worksheets.items) ws.names.load('items/name,items/formula');
        for (const t of workbook.tables.items) t.columns.load('items/name');
        await context.sync();

        const names = [...workbook.names.items, ...workbook.worksheets.items.flatMap((ws) => ws.names.items)]
          .map((n) => ({ name: n.name.slice(n.name.lastIndexOf('!') + 1), refersTo: toText(n.formula) }))
          .filter((n) => !n.name.startsWith('_xl'));

        return {
          sheets: workbook.worksheets.items.map((s) => s.name),
          names,
          tables: workbook.tables.items.map((t) => ({ name: t.name, columns: t.columns.items.map((c) => c.name) })),
        };
      });
    } catch (err) {
      throw translate(err);
    }
  }
}
