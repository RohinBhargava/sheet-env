// Product names the workbook view shows. Ribbon and command names (Home, Insert, Formulas...) are generic
// feature names and are not in here.

export interface Brand {
  /** the product, shown in the title bar and as the title of message boxes */
  excel: string;
  /** the suite name used in "<suite> Scripts" and help text */
  suite: string;
  /** cloud storage names used in the AutoSave tooltip */
  onedrive: string;
  sharepoint: string;
  /** the assistant button label */
  copilot: string;
}

export const BRAND: Brand = {
  excel: 'ForkLaunch Sheets',
  suite: 'ForkLaunch',
  onedrive: 'ForkLaunch cloud storage',
  sharepoint: 'a shared workspace',
  copilot: 'Assistant',
};

/** message-box title, e.g. "ForkLaunch Sheets" */
export const MS_EXCEL = BRAND.excel;
