// Excel ribbon: File, Home, Insert, Page Layout, Formulas, Data, Review, View, Help.
// Commands go to the app through act(cmd, arg); items that are out of scope are drawn as in the real
// app and shown disabled (never a dead click).
import { useEffect, useState, type ReactNode } from 'react';
import { Icon } from './icons';
import { menus, type MenuItem } from '../shell/Menu';
import { BRAND } from '../util/brand';

export type Act = (cmd: string, arg?: unknown) => void;

export interface RibbonState {
  pv: boolean;
  bold: boolean; italic: boolean; underline: boolean;
  fontName: string; fontSize: number;
  hAlign?: string; vAlign?: string; wrap: boolean; merged: boolean;
  numFmtLabel: string;
  showGrid: boolean; showHeadings: boolean; showFormulaBar: boolean; showFormulas: boolean;
  filterOn: boolean; frozen: boolean;
  zoom: number;
  /** the active sheet is protected (Protect Sheet becomes Unprotect Sheet) */
  sheetProtected?: boolean;
}

const ID = {
  home: 'excel.home_tab_number_format_font_alignment_borders',
  formulas: 'excel.formulas_autosum_function_library',
  data: 'excel.data_sort_and_filter',
  insert: 'excel.insert_charts_pivottable_power_query_copilot',
  find: 'excel.find_and_replace_ctrl_f_h',
  goto: 'excel.go_to_ctrl_g_f5',
  copy: 'excel.copy_ctrl_c_marching_ants',
  freeze: 'excel.freeze_panes_merged_cells_rendering',
  save: 'excel.save_save_as_xlsx',
};

export const TABS = ['File', 'Home', 'Insert', 'Draw', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate', 'Add-ins', 'Help'] as const;
export type TabName = typeof TABS[number];

const FONTS = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Calibri Light', 'Cambria', 'Consolas', 'Segoe UI', 'Times New Roman'];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const COLORS = ['#000000', '#FFFFFF', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#C00000', '#FF0000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

function openMenu(el: HTMLElement, items: MenuItem[], controlId: string, minWidth = 180) {
  const r = el.getBoundingClientRect();
  menus.open({ x: r.left, y: r.bottom + 1, items, controlId, minWidth, variant: 'classic', anchor: el, focusFirst: false });
}

function Big({ icon, label, onClick, id, disabled, caret, on, tip, item }: { icon: string; label: ReactNode; onClick?: (el: HTMLElement) => void; id?: string; disabled?: boolean; caret?: boolean; on?: boolean; tip?: string; item?: string }) {
  return (
    <button className={`xl-big${on ? ' on' : ''}`} data-control-id={id} data-control-item={item ?? (typeof label === 'string' ? label.replace(/\n/g, ' ') : undefined)} disabled={disabled} title={tip}
      onMouseDown={e => e.preventDefault()} onClick={e => onClick?.(e.currentTarget)}>
      <Icon name={icon} size={30} />
      <span style={{ whiteSpace: 'pre-line' }}>{label}{caret && <>{typeof label === 'string' && !label.includes('\n') ? <br /> : ' '}<Icon name="chevronDown" size={8} style={{ display: 'inline-block', verticalAlign: 'middle' }} /></>}</span>
    </button>
  );
}
function Btn({ icon, label, onClick, id, disabled, caret, on, tip, item, color }: { icon?: string; label?: string; onClick?: (el: HTMLElement) => void; id?: string; disabled?: boolean; caret?: boolean; on?: boolean; tip?: string; item?: string; color?: string }) {
  return (
    <button className={`xl-btn${on ? ' on' : ''}`} data-control-id={id} data-control-item={item ?? label ?? tip} disabled={disabled} title={tip} aria-label={tip ?? label} aria-pressed={on}
      onMouseDown={e => e.preventDefault()} onClick={e => onClick?.(e.currentTarget)}>
      {icon && (color ? <span className="xl-colorbar"><Icon name={icon} size={16} /><i style={{ background: color }} /></span> : <Icon name={icon} size={16} />)}
      {label && <span>{label}</span>}
      {caret && <Icon name="chevronDown" size={8} />}
    </button>
  );
}
function Group({ label, children, launcher, id, launcherTip }: { label: string; children: ReactNode; launcher?: () => void; id?: string; launcherTip?: string }) {
  return (
    <div className="xl-group" data-control-id={id} data-control-item={label}>
      <div className="xl-group-body">{children}</div>
      <div className="xl-group-label">{label}{launcher !== undefined && <button className="xl-group-launcher" title={launcherTip ?? `${label} Settings`} onClick={launcher} disabled={!launcher}><svg width="8" height="8" viewBox="0 0 8 8"><path d="M1 1h3M1 1v3M2 2l5 5M7 4v3H4" stroke="#555" fill="none" /></svg></button>}</div>
    </div>
  );
}
function Combo({ value, width, items, onPick, id, tip, disabled }: { value: string; width: number; items: string[]; onPick: (v: string) => void; id: string; tip: string; disabled?: boolean }) {
  const [text, setText] = useState(value);
  useEffect(() => setText(value), [value]);
  return (
    <div className="xl-combo" style={{ width, opacity: disabled ? .5 : 1 }} data-control-id={id} data-control-item={tip} title={tip}>
      <input value={text} disabled={disabled} aria-label={tip} style={{ width: width - 18 }} onChange={e => setText(e.target.value)}
        onKeyDown={e => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); onPick(text); } if (e.key === 'Escape') setText(value); }} />
      <button disabled={disabled} aria-label={`${tip} list`} onMouseDown={e => e.preventDefault()} onClick={e => openMenu(e.currentTarget.parentElement!, items.map(v => ({ label: v, checked: v === value, onClick: () => onPick(v) })), id, width)}><Icon name="chevronDown" size={8} /></button>
    </div>
  );
}

export function Ribbon({ tab, st, act, floating, onCollapse }: { tab: TabName; st: RibbonState; act: Act; floating?: boolean; onCollapse?: () => void }) {
  const dis = st.pv; // Protected View: commands are unavailable
  const colorMenu = (el: HTMLElement, cmd: string, auto: string) => openMenu(el, [
    { label: auto, onClick: () => act(cmd, null) },
    { separator: true },
    ...COLORS.map(c => ({ label: c, icon: <span style={{ display: 'inline-block', width: 12, height: 12, background: c, border: '1px solid #bbb' }} />, onClick: () => act(cmd, c) })),
  ], ID.home, 160);
  const fnMenu = (el: HTMLElement, names: string[]) => openMenu(el, [...names.map(n => ({ label: n, onClick: () => act('insertFunction', n) })), { separator: true }, { label: 'Insert Function...', onClick: () => act('insertFunction', '') }], ID.formulas, 160);
  let body: ReactNode = null;
  switch (tab) {
    case 'Home':
      body = (
        <>
          <Group label="Clipboard" launcher={undefined}>
            <div className="xl-split" style={{ flexDirection: 'column' }}>
              <Big icon="paste" label="Paste" id={ID.copy} item="Paste" disabled={dis} onClick={() => act('paste')} />
            </div>
            <div className="xl-col">
              <Btn icon="cut" tip="Cut (Ctrl+X)" id={ID.copy} item="Cut" disabled={dis} onClick={() => act('cut')} />
              <Btn icon="copy" tip="Copy (Ctrl+C)" id={ID.copy} item="Copy" disabled={dis} caret onClick={() => act('copy')} />
              <Btn icon="painter" tip="Format Painter" id={ID.home} item="Format Painter" disabled={dis} onClick={() => act('formatPainter')} />
            </div>
          </Group>
          <Group label="Font" launcher={dis ? undefined : () => act('formatCells', 'font')} id={ID.home}>
            <div className="xl-col">
              <div className="xl-row">
                <Combo value={st.fontName} width={124} items={FONTS} onPick={v => act('fontName', v)} id={ID.home} tip="Font" disabled={dis} />
                <Combo value={String(st.fontSize)} width={46} items={SIZES.map(String)} onPick={v => act('fontSize', Number(v))} id={ID.home} tip="Font Size" disabled={dis} />
                <Btn icon="grow" tip="Increase Font Size" id={ID.home} disabled={dis} onClick={() => act('fontGrow', 1)} />
                <Btn icon="shrink" tip="Decrease Font Size" id={ID.home} disabled={dis} onClick={() => act('fontGrow', -1)} />
              </div>
              <div className="xl-row">
                <Btn icon="bold" tip="Bold (Ctrl+B)" id={ID.home} item="Bold" on={st.bold} disabled={dis} onClick={() => act('bold')} />
                <Btn icon="italic" tip="Italic (Ctrl+I)" id={ID.home} item="Italic" on={st.italic} disabled={dis} onClick={() => act('italic')} />
                <Btn icon="underline" tip="Underline (Ctrl+U)" id={ID.home} item="Underline" on={st.underline} disabled={dis} onClick={() => act('underline')} />
                <Btn icon="borders" tip="Borders" id={ID.home} item="Borders" caret disabled={dis} onClick={el => openMenu(el, [
                  { label: 'Bottom Border', onClick: () => act('border', 'bottom') }, { label: 'Top Border', onClick: () => act('border', 'top') },
                  { label: 'Left Border', onClick: () => act('border', 'left') }, { label: 'Right Border', onClick: () => act('border', 'right') },
                  { separator: true }, { label: 'No Border', onClick: () => act('border', 'none') }, { label: 'All Borders', onClick: () => act('border', 'all') },
                  { label: 'Outside Borders', onClick: () => act('border', 'outside') }, { label: 'Thick Bottom Border', onClick: () => act('border', 'thickBottom') },
                ], ID.home, 200)} />
                <Btn icon="fill" tip="Fill Color" id={ID.home} item="Fill Color" caret color="#FFFF00" disabled={dis} onClick={el => colorMenu(el, 'fill', 'No Fill')} />
                <Btn icon="fontColor" tip="Font Color" id={ID.home} item="Font Color" caret color="#FF0000" disabled={dis} onClick={el => colorMenu(el, 'fontColor', 'Automatic')} />
              </div>
            </div>
          </Group>
          <Group label="Alignment" launcher={dis ? undefined : () => act('formatCells', 'alignment')} id={ID.home}>
            <div className="xl-col">
              <div className="xl-row">
                <Btn icon="alignTop" tip="Top Align" id={ID.home} on={st.vAlign === 'top'} disabled={dis} onClick={() => act('vAlign', 'top')} />
                <Btn icon="alignMiddle" tip="Middle Align" id={ID.home} on={st.vAlign === 'center'} disabled={dis} onClick={() => act('vAlign', 'center')} />
                <Btn icon="alignBottom" tip="Bottom Align" id={ID.home} on={!st.vAlign || st.vAlign === 'bottom'} disabled={dis} onClick={() => act('vAlign', 'bottom')} />
                <Btn icon="orientation" tip="Orientation" id={ID.home} caret disabled={dis} onClick={el => openMenu(el, [{ label: 'Angle Counterclockwise', onClick: () => act('rotate', 45) }, { label: 'Angle Clockwise', onClick: () => act('rotate', 135) }, { label: 'Rotate Text Up', onClick: () => act('rotate', 90) }, { label: 'Rotate Text Down', onClick: () => act('rotate', 180) }, { label: 'Format Cell Alignment', onClick: () => act('formatCells', 'alignment') }], ID.home)} />
                <Btn icon="wrap" label="Wrap Text" id={ID.home} on={st.wrap} disabled={dis} onClick={() => act('wrap')} />
              </div>
              <div className="xl-row">
                <Btn icon="alignLeft" tip="Align Left" id={ID.home} on={st.hAlign === 'left'} disabled={dis} onClick={() => act('hAlign', 'left')} />
                <Btn icon="alignCenter" tip="Center" id={ID.home} on={st.hAlign === 'center'} disabled={dis} onClick={() => act('hAlign', 'center')} />
                <Btn icon="alignRight" tip="Align Right" id={ID.home} on={st.hAlign === 'right'} disabled={dis} onClick={() => act('hAlign', 'right')} />
                <Btn icon="indentDec" tip="Decrease Indent" id={ID.home} disabled={dis} onClick={() => act('indent', -1)} />
                <Btn icon="indentInc" tip="Increase Indent" id={ID.home} disabled={dis} onClick={() => act('indent', 1)} />
                <div className="xl-split">
                  <Btn icon="merge" label="Merge & Center" id={ID.home} item="Merge & Center" on={st.merged} disabled={dis} onClick={() => act('mergeCenter')} />
                  <Btn tip="Merge options" id={ID.home} caret disabled={dis} onClick={el => openMenu(el, [{ label: 'Merge & Center', onClick: () => act('mergeCenter') }, { label: 'Merge Across', onClick: () => act('mergeAcross') }, { label: 'Merge Cells', onClick: () => act('mergeCells') }, { label: 'Unmerge Cells', onClick: () => act('unmerge') }], ID.home)} />
                </div>
              </div>
            </div>
          </Group>
          <Group label="Number" launcher={dis ? undefined : () => act('formatCells', 'number')} id={ID.home}>
            <div className="xl-col">
              <div className="xl-row">
                <Combo value={st.numFmtLabel} width={128} items={['General', 'Number', 'Currency', 'Accounting', 'Short Date', 'Long Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text']} onPick={v => act('numFmt', v)} id={ID.home} tip="Number Format" disabled={dis} />
              </div>
              <div className="xl-row">
                <Btn icon="currency" tip="Accounting Number Format" id={ID.home} caret disabled={dis} onClick={() => act('numFmt', 'Accounting')} />
                <Btn icon="percent" tip="Percent Style (Ctrl+Shift+%)" id={ID.home} disabled={dis} onClick={() => act('numFmt', 'Percentage0')} />
                <Btn icon="comma" tip="Comma Style" id={ID.home} disabled={dis} onClick={() => act('numFmt', 'Comma')} />
                <Btn icon="decInc" tip="Increase Decimal" id={ID.home} disabled={dis} onClick={() => act('decimals', 1)} />
                <Btn icon="decDec" tip="Decrease Decimal" id={ID.home} disabled={dis} onClick={() => act('decimals', -1)} />
              </div>
            </div>
          </Group>
          <Group label="Styles">
            <span className="xl-wide" style={{ display: 'contents' }}>
              <Big icon="condFormat" label={'Conditional\nFormatting'} caret disabled />
              <Big icon="formatTable" label={'Format as\nTable'} caret disabled />
              <Big icon="cellStyles" label={'Cell\nStyles'} caret disabled />
            </span>
            <span className="xl-narrow"><Big icon="cellStyles" label="Styles" caret disabled /></span>
          </Group>
          <Group label="Cells">
            <Big icon="insertCells" label="Insert" caret disabled={dis} id={ID.home} onClick={el => openMenu(el, [{ label: 'Insert Cells...', disabled: true }, { label: 'Insert Sheet Rows', onClick: () => act('insertRows') }, { label: 'Insert Sheet Columns', onClick: () => act('insertCols') }, { label: 'Insert Sheet', onClick: () => act('insertSheet') }], ID.home)} />
            <Big icon="deleteCells" label="Delete" caret disabled={dis} id={ID.home} onClick={el => openMenu(el, [{ label: 'Delete Cells...', disabled: true }, { label: 'Delete Sheet Rows', onClick: () => act('deleteRows') }, { label: 'Delete Sheet Columns', onClick: () => act('deleteCols') }, { label: 'Delete Sheet', onClick: () => act('deleteSheet') }], ID.home)} />
            <Big icon="format" label="Format" caret disabled={dis} id={ID.home} onClick={el => openMenu(el, [
              { label: 'Row Height...', onClick: () => act('rowHeightDialog') }, { label: 'AutoFit Row Height', onClick: () => act('autofitRows') },
              { label: 'Column Width...', onClick: () => act('colWidthDialog') }, { label: 'AutoFit Column Width', onClick: () => act('autofitCols') },
              { label: 'Default Width...', disabled: true }, { separator: true },
              { label: 'Hide & Unhide', submenu: [{ label: 'Hide Rows', onClick: () => act('hideRows') }, { label: 'Hide Columns', onClick: () => act('hideCols') }, { label: 'Unhide Rows', onClick: () => act('unhideRows') }, { label: 'Unhide Columns', onClick: () => act('unhideCols') }] },
              { separator: true }, { label: 'Rename Sheet', onClick: () => act('renameSheet') }, { label: 'Move or Copy Sheet...', onClick: () => act('moveCopySheet') },
              { separator: true }, { label: st.sheetProtected ? 'Unprotect Sheet...' : 'Protect Sheet...', onClick: () => act('protectSheet') }, { label: 'Format Cells...', onClick: () => act('formatCells', 'number') },
            ], ID.home, 200)} />
          </Group>
          <Group label="Editing">
            <div className="xl-col">
              <Btn icon="autosum" label="AutoSum" caret id={ID.formulas} item="AutoSum" disabled={dis} onClick={() => act('autosum', 'SUM')} />
              <Btn icon="fillDown" label="Fill" caret id={ID.home} disabled={dis} onClick={el => openMenu(el, [{ label: 'Down', onClick: () => act('fill', 'down') }, { label: 'Right', onClick: () => act('fill', 'right') }, { label: 'Up', onClick: () => act('fill', 'up') }, { label: 'Left', onClick: () => act('fill', 'left') }], ID.home)} />
              <Btn icon="clear" label="Clear" caret id={ID.home} disabled={dis} onClick={el => openMenu(el, [{ label: 'Clear All', onClick: () => act('clear', 'all') }, { label: 'Clear Formats', onClick: () => act('clear', 'formats') }, { label: 'Clear Contents', onClick: () => act('clear', 'contents') }], ID.home)} />
            </div>
            <Big icon="sort" label={'Sort &\nFilter'} caret id={ID.data} item="Sort & Filter" disabled={dis} onClick={el => openMenu(el, [{ label: 'Sort A to Z', onClick: () => act('sort', 1) }, { label: 'Sort Z to A', onClick: () => act('sort', -1) }, { label: 'Custom Sort...', disabled: true }, { separator: true }, { label: 'Filter', checked: st.filterOn, onClick: () => act('filter') }, { label: 'Clear', disabled: !st.filterOn, onClick: () => act('filter') }, { label: 'Reapply', disabled: true }], ID.data)} />
            <Big icon="find" label={'Find &\nSelect'} caret id={ID.find} item="Find & Select" onClick={el => openMenu(el, [
              { label: 'Find...', shortcut: 'Ctrl+F', controlId: ID.find, onClick: () => act('find') },
              { label: 'Replace...', shortcut: 'Ctrl+H', controlId: ID.find, disabled: dis, onClick: () => act('replace') },
              { label: 'Go To...', shortcut: 'Ctrl+G', controlId: ID.goto, onClick: () => act('goto') },
              { label: 'Go To Special...', controlId: ID.goto, disabled: true },
              { separator: true }, { label: 'Formulas', onClick: () => act('selectSpecial', 'formulas') }, { label: 'Constants', onClick: () => act('selectSpecial', 'constants') },
            ], ID.find, 200)} />
          </Group>
        </>
      );
      break;
    case 'Insert':
      body = (
        <>
          <Group label="Tables"><Big icon="pivot" label="PivotTable" caret id={ID.insert} disabled /><Big icon="pivot" label={'Recommended\nPivotTables'} id={ID.insert} disabled /><Big icon="table" label="Table" id={ID.insert} disabled /></Group>
          <Group label="Illustrations"><Big icon="picture" label="Pictures" caret id={ID.insert} disabled /><Big icon="shapes" label="Shapes" caret id={ID.insert} disabled /></Group>
          <Group label="Charts" launcher={undefined}><Big icon="chart" label={'Recommended\nCharts'} id={ID.insert} disabled /><div className="xl-col"><Btn icon="chart" tip="Insert Column or Bar Chart" caret id={ID.insert} disabled /><Btn icon="line" tip="Insert Line or Area Chart" caret id={ID.insert} disabled /><Btn icon="pie" tip="Insert Pie or Doughnut Chart" caret id={ID.insert} disabled /></div><Big icon="pivot" label="PivotChart" caret id={ID.insert} disabled /></Group>
          <Group label="Sparklines"><Big icon="line" label="Line" id={ID.insert} disabled /><Big icon="chart" label="Column" id={ID.insert} disabled /></Group>
          <Group label="Filters"><Big icon="filter" label="Slicer" id={ID.insert} disabled /><Big icon="clock" label="Timeline" id={ID.insert} disabled /></Group>
          <Group label="Links"><Big icon="link" label="Link" caret id={ID.insert} disabled /></Group>
          <Group label="Comments"><Big icon="comment" label="Comment" id={ID.insert} disabled /></Group>
          <Group label="Text"><Big icon="textbox" label="Text" caret id={ID.insert} disabled /></Group>
          <Group label="Symbols"><Big icon="symbol" label="Symbols" caret id={ID.insert} disabled /></Group>
        </>
      );
      break;
    case 'Page Layout':
      body = (
        <>
          <Group label="Themes"><Big icon="themes" label="Themes" caret disabled /></Group>
          <Group label="Page Setup" launcher={undefined}><Big icon="margins" label="Margins" caret disabled /><Big icon="orientationPage" label="Orientation" caret disabled /><Big icon="size" label="Size" caret disabled /><Big icon="printArea" label={'Print\nArea'} caret disabled /><Big icon="breaks" label="Breaks" caret disabled /><Big icon="background" label="Background" disabled /><Big icon="printTitles" label={'Print\nTitles'} disabled /></Group>
          <Group label="Scale to Fit"><div className="xl-col"><div className="xl-row" style={{ gap: 6 }}><span style={{ width: 42 }}>Width:</span><div className="xl-combo" style={{ width: 90, opacity: .5 }}><input value="Automatic" disabled readOnly style={{ width: 70 }} /></div></div><div className="xl-row" style={{ gap: 6 }}><span style={{ width: 42 }}>Height:</span><div className="xl-combo" style={{ width: 90, opacity: .5 }}><input value="Automatic" disabled readOnly style={{ width: 70 }} /></div></div><div className="xl-row" style={{ gap: 6 }}><span style={{ width: 42 }}>Scale:</span><div className="xl-combo" style={{ width: 90, opacity: .5 }}><input value="100%" disabled readOnly style={{ width: 70 }} /></div></div></div></Group>
          <Group label="Sheet Options">
            <div className="xl-col" style={{ paddingTop: 2 }}>
              <div className="xl-row" style={{ gap: 12 }}><b style={{ fontWeight: 400, width: 60 }}>Gridlines</b><b style={{ fontWeight: 400 }}>Headings</b></div>
              <div className="xl-row" style={{ gap: 12 }}>
                <label className="xl-check" style={{ width: 60 }}><input type="checkbox" checked={st.showGrid} onChange={() => act('toggleGrid')} data-control-id={ID.home} data-control-item="View Gridlines" />View</label>
                <label className="xl-check"><input type="checkbox" checked={st.showHeadings} onChange={() => act('toggleHeadings')} />View</label>
              </div>
              <div className="xl-row" style={{ gap: 12 }}>
                <label className="xl-check" style={{ width: 60, opacity: .5 }}><input type="checkbox" disabled />Print</label>
                <label className="xl-check" style={{ opacity: .5 }}><input type="checkbox" disabled />Print</label>
              </div>
            </div>
          </Group>
          <Group label="Arrange"><Big icon="arrange" label={'Bring\nForward'} disabled /><Big icon="arrange" label={'Send\nBackward'} disabled /><Big icon="arrange" label={'Selection\nPane'} disabled /></Group>
        </>
      );
      break;
    case 'Formulas':
      body = (
        <>
          <Group label="Function Library">
            <Big icon="function" label={'Insert\nFunction'} id={ID.formulas} item="Insert Function" disabled={dis} onClick={() => act('insertFunction', '')} />
            <Big icon="autosum" label="AutoSum" caret id={ID.formulas} item="AutoSum" disabled={dis} onClick={el => openMenu(el, [{ label: 'Sum', onClick: () => act('autosum', 'SUM') }, { label: 'Average', onClick: () => act('autosum', 'AVERAGE') }, { label: 'Count Numbers', onClick: () => act('autosum', 'COUNT') }, { label: 'Max', onClick: () => act('autosum', 'MAX') }, { label: 'Min', onClick: () => act('autosum', 'MIN') }], ID.formulas)} />
            <Big icon="clock" label={'Recently\nUsed'} caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV'])} />
            <Big icon="calc" label="Financial" caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['FV', 'NPV', 'PMT', 'PV'])} />
            <Big icon="logical" label="Logical" caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['AND', 'FALSE', 'IF', 'IFERROR', 'IFNA', 'IFS', 'NOT', 'OR', 'SWITCH', 'TRUE', 'XOR'])} />
            <Big icon="text" label="Text" caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['CHAR', 'CLEAN', 'CODE', 'CONCAT', 'CONCATENATE', 'DOLLAR', 'EXACT', 'FIND', 'FIXED', 'LEFT', 'LEN', 'LOWER', 'MID', 'PROPER', 'REPLACE', 'REPT', 'RIGHT', 'SEARCH', 'SUBSTITUTE', 'T', 'TEXT', 'TEXTJOIN', 'TRIM', 'UPPER', 'VALUE'])} />
            <Big icon="clock" label={'Date &\nTime'} caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'EDATE', 'EOMONTH', 'HOUR', 'MINUTE', 'MONTH', 'NETWORKDAYS', 'NOW', 'SECOND', 'TIME', 'TODAY', 'WEEKDAY', 'YEAR', 'YEARFRAC'])} />
            <Big icon="lookup" label={'Lookup &\nReference'} caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['CHOOSE', 'COLUMN', 'COLUMNS', 'HLOOKUP', 'HYPERLINK', 'INDEX', 'MATCH', 'ROW', 'ROWS', 'TRANSPOSE', 'UNIQUE', 'VLOOKUP', 'XLOOKUP', 'XMATCH'])} />
            <Big icon="math" label={'Math &\nTrig'} caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['ABS', 'CEILING', 'EXP', 'FLOOR', 'INT', 'LN', 'LOG', 'LOG10', 'MOD', 'MROUND', 'PI', 'POWER', 'PRODUCT', 'RAND', 'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'SIGN', 'SQRT', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'TRUNC'])} />
            <Big icon="more" label={'More\nFunctions'} caret id={ID.formulas} disabled={dis} onClick={el => fnMenu(el, ['AVERAGEIF', 'AVERAGEIFS', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'LARGE', 'MAXIFS', 'MEDIAN', 'MINIFS', 'RANK.EQ', 'SMALL', 'STDEV.S'])} />
          </Group>
          <Group label="Defined Names"><Big icon="names" label={'Name\nManager'} disabled /><div className="xl-col"><Btn icon="names" label="Define Name" caret disabled /><Btn icon="function" label="Use in Formula" caret disabled /><Btn icon="table" label="Create from Selection" disabled /></div></Group>
          <Group label="Formula Auditing"><div className="xl-col"><Btn icon="trace" label="Trace Precedents" disabled /><Btn icon="trace" label="Trace Dependents" disabled /><Btn icon="clear" label="Remove Arrows" caret disabled /></div><div className="xl-col"><Btn icon="showFormulas" label="Show Formulas" id={ID.formulas} on={st.showFormulas} onClick={() => act('toggleFormulas')} /><Btn icon="validation" label="Error Checking" caret disabled /><Btn icon="function" label="Evaluate Formula" disabled /></div><Big icon="zoom" label={'Watch\nWindow'} disabled /></Group>
          <Group label="Calculation"><Big icon="calcOptions" label={'Calculation\nOptions'} caret id={ID.formulas} onClick={el => openMenu(el, [{ label: 'Automatic', checked: true, onClick: () => {} }, { label: 'Automatic Except for Data Tables', onClick: () => {} }, { label: 'Manual', disabled: true }], ID.formulas)} /><div className="xl-col"><Btn icon="calc" label="Calculate Now" id={ID.formulas} onClick={() => act('calculate')} /><Btn icon="calc" label="Calculate Sheet" id={ID.formulas} onClick={() => act('calculate')} /></div></Group>
        </>
      );
      break;
    case 'Data':
      body = (
        <>
          <Group label="Get & Transform Data"><Big icon="getData" label={'Get\nData'} caret id={ID.insert} disabled /><div className="xl-col"><Btn icon="table" label="From Text/CSV" id={ID.insert} disabled /><Btn icon="link" label="From Web" id={ID.insert} disabled /><Btn icon="table" label="From Table/Range" id={ID.insert} disabled /></div></Group>
          <Group label="Queries & Connections"><Big icon="refresh" label={'Refresh\nAll'} caret disabled /></Group>
          <Group label="Sort & Filter">
            <div className="xl-col"><Btn icon="sort" tip="Sort A to Z" id={ID.data} item="Sort A to Z" disabled={dis} onClick={() => act('sort', 1)} /><Btn icon="sortZA" tip="Sort Z to A" id={ID.data} item="Sort Z to A" disabled={dis} onClick={() => act('sort', -1)} /></div>
            <Big icon="sort" label="Sort" id={ID.data} disabled={dis} onClick={() => act('sort', 1)} />
            <Big icon="filter" label="Filter" id={ID.data} on={st.filterOn} disabled={dis} onClick={() => act('filter')} />
            <div className="xl-col"><Btn icon="clear" label="Clear" id={ID.data} disabled={!st.filterOn} onClick={() => act('filter')} /><Btn icon="refresh" label="Reapply" disabled /><Btn icon="filter" label="Advanced" disabled /></div>
          </Group>
          <Group label="Data Tools"><Big icon="textToCols" label={'Text to\nColumns'} disabled /><div className="xl-col"><Btn icon="fillDown" tip="Flash Fill" disabled /><Btn icon="dupes" tip="Remove Duplicates" disabled /><Btn icon="validation" tip="Data Validation" caret disabled /></div></Group>
          <Group label="Forecast"><Big icon="whatif" label={'What-If\nAnalysis'} caret disabled /><Big icon="line" label={'Forecast\nSheet'} disabled /></Group>
          <Group label="Outline"><Big icon="group" label="Group" caret disabled /><Big icon="group" label="Ungroup" caret disabled /><Big icon="autosum" label="Subtotal" disabled /></Group>
        </>
      );
      break;
    case 'Review':
      body = (
        <>
          <Group label="Proofing"><Big icon="spelling" label="Spelling" id={ID.home} item="Spelling" onClick={() => act('spelling')} /><Big icon="thesaurus" label="Thesaurus" disabled /></Group>
          <Group label="Accessibility"><Big icon="access" label={'Check\nAccessibility'} caret onClick={() => act('accessibility')} /></Group>
          <Group label="Language"><Big icon="text" label="Translate" disabled /></Group>
          <Group label="Comments"><Big icon="comment" label={'New\nComment'} disabled /><div className="xl-col"><Btn icon="deleteCells" label="Delete" disabled /><Btn icon="chevronLeft" label="Previous Comment" disabled /><Btn icon="chevronRight" label="Next Comment" disabled /></div><Big icon="comment" label={'Show\nComments'} disabled /></Group>
          <Group label="Notes"><Big icon="note" label="Notes" caret disabled /></Group>
          <Group label="Protect"><Big icon="protect" label={st.sheetProtected ? 'Unprotect\nSheet' : 'Protect\nSheet'} disabled={dis} id="excel.sheet_tabs" onClick={() => act('protectSheet')} /><Big icon="protect" label={'Protect\nWorkbook'} disabled /><Big icon="protect" label={'Allow Edit\nRanges'} disabled /></Group>
        </>
      );
      break;
    case 'View':
      body = (
        <>
          <Group label="Sheet View"><Big icon="normalView" label={'Sheet\nView'} caret disabled /></Group>
          <Group label="Workbook Views"><Big icon="normalView" label="Normal" on onClick={() => {}} /><Big icon="pageBreakView" label={'Page Break\nPreview'} disabled /><Big icon="pageLayoutView" label={'Page\nLayout'} disabled /><Big icon="normalView" label={'Custom\nViews'} disabled /></Group>
          <Group label="Show">
            <div className="xl-col" style={{ paddingTop: 4 }}>
              <label className="xl-check"><input type="checkbox" disabled />Ruler</label>
              <label className="xl-check"><input type="checkbox" checked={st.showGrid} onChange={() => act('toggleGrid')} />Gridlines</label>
            </div>
            <div className="xl-col" style={{ paddingTop: 4 }}>
              <label className="xl-check"><input type="checkbox" checked={st.showFormulaBar} onChange={() => act('toggleFormulaBar')} />Formula Bar</label>
              <label className="xl-check"><input type="checkbox" checked={st.showHeadings} onChange={() => act('toggleHeadings')} />Headings</label>
            </div>
          </Group>
          <Group label="Zoom"><Big icon="zoom" label="Zoom" onClick={() => act('zoomDialog')} /><Big icon="zoom100" label="100%" onClick={() => act('zoom', 1)} /><Big icon="zoomSel" label={'Zoom to\nSelection'} onClick={() => act('zoomSel')} /></Group>
          <Group label="Window">
            <Big icon="newWindow" label={'New\nWindow'} disabled /><Big icon="arrange" label={'Arrange\nAll'} disabled />
            <Big icon="freeze" label={'Freeze\nPanes'} caret id={ID.freeze} item="Freeze Panes" onClick={el => openMenu(el, [
              { label: st.frozen ? 'Unfreeze Panes' : 'Freeze Panes', controlId: ID.freeze, onClick: () => act(st.frozen ? 'unfreeze' : 'freeze') },
              { label: 'Freeze Top Row', controlId: ID.freeze, onClick: () => act('freezeTopRow') },
              { label: 'Freeze First Column', controlId: ID.freeze, onClick: () => act('freezeFirstCol') },
            ], ID.freeze, 220)} />
            <div className="xl-col"><Btn icon="arrange" tip="Split" disabled /><Btn icon="arrange" tip="Hide" disabled /><Btn icon="arrange" tip="Unhide" disabled /></div>
            <Big icon="switch" label={'Switch\nWindows'} caret onClick={el => act('switchWindows', el)} />
          </Group>
          <Group label="Macros"><Big icon="macros" label="Macros" caret disabled /></Group>
        </>
      );
      break;
    case 'Draw':
      body = (
        <>
          <Group label="Drawing Tools"><Big icon="shapes" label="Select" disabled /><Big icon="fontColor" label="Pen" disabled /><Big icon="fill" label="Highlighter" disabled /><Big icon="clear" label="Eraser" disabled /></Group>
          <Group label="Convert"><Big icon="shapes" label={'Ink to\nShape'} disabled /><Big icon="math" label={'Ink to\nMath'} disabled /></Group>
        </>
      );
      break;
    case 'Automate':
      body = (
        <>
          <Group label={`${BRAND.suite} Scripts`}><Big icon="macros" label={'New\nScript'} disabled /><Big icon="macros" label={'View\nScripts'} disabled /></Group>
        </>
      );
      break;
    case 'Add-ins':
      body = (
        <>
          <Group label="Add-ins"><Big icon="new" label={'Get\nAdd-ins'} disabled /><Big icon="new" label={'My\nAdd-ins'} disabled /></Group>
        </>
      );
      break;
    case 'Help':
      body = (
        <>
          <Group label="Help"><Big icon="help" label="Help" onClick={() => act('help')} /><Big icon="help" label={'Contact\nSupport'} disabled /><Big icon="feedback" label="Feedback" disabled /><Big icon="training" label={'Show\nTraining'} disabled /><Big icon="info" label={"What's\nNew"} disabled /></Group>
        </>
      );
      break;
  }
  return (
    <div className={`xl-ribbon${floating ? ' floating' : ''}`} role="toolbar" aria-label={`${tab} ribbon`}>
      {body}
      {onCollapse && <button className="xl-ribbon-collapse" title="Collapse the Ribbon (Ctrl+F1)" onClick={onCollapse}><Icon name="chevronUp" size={12} /></button>}
    </div>
  );
}
