import type { WorkbookInfo } from '../editor/context';
import type { CellTarget, SelectionInfo, WorkbookAdapter } from './types';

/** A pretend workbook so the editor can be tried in an ordinary browser, outside Excel. */
export const DEMO_WORKBOOK: WorkbookInfo = {
  sheets: ['Sales', 'Lookup', 'Q1 Targets'],
  names: [
    { name: 'TaxRate', refersTo: '=Lookup!$B$2' },
    { name: 'Revenue', refersTo: '=Sales!$D$2:$D$500' },
    { name: 'MarkUp', refersTo: '=LAMBDA(price, [pct], price * (1 + IF(ISOMITTED(pct), 0.2, pct)))' },
  ],
  tables: [
    { name: 'Sales', columns: ['Region', 'Product', 'Units', 'Revenue', 'Order Date'] },
    { name: 'Targets', columns: ['Region', 'Quarter', 'Target'] },
  ],
};

const DEMO_FORMULA =
  '=LET(rows,FILTER(Sales[[Region]:[Revenue]],(Sales[Region]="West")*(Sales[Units]>10),"none"),total,SUM(INDEX(rows,,4)),IF(total>TaxRate*1000,TEXT(total,"#,##0"),"low"))';

export class StandaloneAdapter implements WorkbookAdapter {
  readonly kind = 'standalone' as const;
  private text = DEMO_FORMULA;
  private target: CellTarget = { sheet: 'Sales', address: 'F2' };

  async getSelection(): Promise<SelectionInfo> {
    return { target: this.target, text: this.text, display: '', cellCount: 1, isFormula: this.text.startsWith('=') };
  }
  async onSelectionChanged(): Promise<void> {}
  async write(_target: CellTarget, text: string): Promise<string> {
    this.text = text;
    return '(preview only)';
  }
  async getWorkbookInfo(): Promise<WorkbookInfo> {
    return DEMO_WORKBOOK;
  }
}
