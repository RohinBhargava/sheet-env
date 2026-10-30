// Our own drawn glyphs for the workbook view (no third-party icon files or icon fonts).
import type { CSSProperties } from 'react';

const G = '#107c41';
const B = '#2b579a';

export function ExcelAppIcon({ size = 24 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 32 32" aria-hidden style={{ display: 'block' }}>
      <rect x="3" y="3" width="26" height="26" rx="5" fill="#1a7f4b" />
      <rect x="7" y="8" width="18" height="17" rx="1.5" fill="#fff" />
      <rect x="7" y="8" width="18" height="4.5" rx="1.5" fill="#bfe6cf" />
      <path d="M7 12.5h18M7 16.7h18M7 20.9h18M13 8v17M19 8v17" stroke="#1a7f4b" strokeWidth="1.2" />
    </svg>
  );
}

type Name = string;

/** stroke glyph set, 16x16 grid unless noted */
export function Icon({ name, size = 16, style, color }: { name: Name; size?: number; style?: CSSProperties; color?: string }) {
  const s = { fill: 'none', stroke: color ?? 'currentColor', strokeWidth: 1, strokeLinecap: 'round' as const, strokeLinejoin: 'round' as const };
  const body = (() => {
    switch (name) {
      case 'save': return <><path d="M2.5 2.5h9l2 2v9h-11z" {...s} /><path d="M5 2.5v3.5h5V2.5M4.5 13.5v-4h7v4" {...s} /></>;
      case 'undo': return <><path d="M5 4L2.5 6.5 5 9" {...s} /><path d="M2.5 6.5H10a3.5 3.5 0 0 1 0 7H7" {...s} /></>;
      case 'redo': return <><path d="M11 4l2.5 2.5L11 9" {...s} /><path d="M13.5 6.5H6a3.5 3.5 0 0 0 0 7h3" {...s} /></>;
      case 'chevronDown': return <path d="M4 6l4 4 4-4" {...s} />;
      case 'chevronRight': return <path d="M6 4l4 4-4 4" {...s} />;
      case 'chevronLeft': return <path d="M10 4L6 8l4 4" {...s} />;
      case 'chevronUp': return <path d="M4 10l4-4 4 4" {...s} />;
      case 'close': return <path d="M3.5 3.5l9 9M12.5 3.5l-9 9" {...s} />;
      case 'check': return <path d="M3 8.5l3 3 7-7" {...s} />;
      case 'search': return <><circle cx="6.5" cy="6.5" r="4" {...s} /><path d="M9.5 9.5l4 4" {...s} /></>;
      case 'shield': return <><path d="M8 1.5l5.5 2v4c0 3.5-2.4 5.9-5.5 7-3.1-1.1-5.5-3.5-5.5-7v-4z" fill="#fff" stroke="#8a6d00" /><path d="M8 4.5v4.5M8 11v.5" stroke="#8a6d00" strokeLinecap="round" /></>;
      case 'paste': return <><rect x="4" y="3" width="8" height="11" rx="1" {...s} /><path d="M6 3V2h4v1" {...s} /><rect x="7" y="7" width="7" height="8" fill="#fff" stroke={B} /><path d="M8.5 9.5h4M8.5 11.5h4M8.5 13.5h2.5" stroke={B} /></>;
      case 'cut': return <><circle cx="4.5" cy="12" r="2" {...s} /><circle cx="11.5" cy="12" r="2" {...s} /><path d="M6 10.5L12 2M10 10.5L4 2" {...s} /></>;
      case 'copy': return <><rect x="5.5" y="4.5" width="8" height="9.5" rx="1" {...s} /><path d="M3.5 11.5V2.5h7" {...s} /></>;
      case 'painter': return <><rect x="3" y="2" width="9" height="4" rx="1" {...s} /><path d="M12 4h1.5v3.5H8V9" {...s} /><rect x="7" y="9" width="2" height="5" rx=".5" {...s} /></>;
      case 'bold': return <path d="M5 3h4a2.5 2.5 0 0 1 0 5H5zM5 8h4.5a2.8 2.8 0 0 1 0 5.5H5z" fill="none" stroke="currentColor" strokeWidth="1.8" />;
      case 'italic': return <path d="M7 3h5M4 13h5M9.5 3l-3 10" {...s} strokeWidth={1.3} />;
      case 'underline': return <><path d="M5 2.5v5a3 3 0 0 0 6 0v-5" {...s} /><path d="M3.5 14h9" {...s} /></>;
      case 'borders': return <><rect x="2.5" y="2.5" width="11" height="11" stroke="currentColor" strokeDasharray="1 1.5" fill="none" /><path d="M2.5 13.5h11" stroke="currentColor" strokeWidth="1.5" /></>;
      case 'fill': return <><path d="M3 8.5L7.5 4l4.5 4.5L7.5 13z" {...s} /><path d="M12.5 9.5s1.2 1.6 1.2 2.4a1.2 1.2 0 0 1-2.4 0c0-.8 1.2-2.4 1.2-2.4z" fill="currentColor" /></>;
      case 'fontColor': return <path d="M4 12L8 2.5 12 12M5.5 8.5h5" {...s} />;
      case 'grow': return <><path d="M2 13L5.5 4 9 13M3.3 10h4.4" {...s} /><path d="M11 6l2-2.5L15 6" {...s} /></>;
      case 'shrink': return <><path d="M2 13L5.5 4 9 13M3.3 10h4.4" {...s} /><path d="M11 4l2 2.5L15 4" {...s} /></>;
      case 'alignTop': return <><path d="M2.5 2.5h11" {...s} /><path d="M8 13V5M5 7.5l3-3 3 3" {...s} /></>;
      case 'alignMiddle': return <><path d="M2.5 8h11" {...s} /><path d="M8 2v3.5M8 10.5V14M6 4l2 2 2-2M6 12l2-2 2 2" {...s} /></>;
      case 'alignBottom': return <><path d="M2.5 13.5h11" {...s} /><path d="M8 3v8M5 8.5l3 3 3-3" {...s} /></>;
      case 'alignLeft': return <path d="M2.5 3.5h11M2.5 6.5h7M2.5 9.5h11M2.5 12.5h7" {...s} />;
      case 'alignCenter': return <path d="M2.5 3.5h11M4.5 6.5h7M2.5 9.5h11M4.5 12.5h7" {...s} />;
      case 'alignRight': return <path d="M2.5 3.5h11M6.5 6.5h7M2.5 9.5h11M6.5 12.5h7" {...s} />;
      case 'wrap': return <><path d="M2.5 3.5h11M2.5 12.5h4" {...s} /><path d="M2.5 8h9a2 2 0 0 1 0 4.5H9M10.5 11l-1.5 1.5 1.5 1.5" {...s} /></>;
      case 'merge': return <><rect x="1.5" y="3.5" width="13" height="9" {...s} /><path d="M4 8h8M5.5 6.5L4 8l1.5 1.5M10.5 6.5L12 8l-1.5 1.5" {...s} /></>;
      case 'indentDec': return <><path d="M6.5 3.5h7M6.5 6.5h7M6.5 9.5h7M2.5 12.5h11" {...s} /><path d="M4.5 5L2.5 7l2 2" {...s} /></>;
      case 'indentInc': return <><path d="M6.5 3.5h7M6.5 6.5h7M6.5 9.5h7M2.5 12.5h11" {...s} /><path d="M2.5 5l2 2-2 2" {...s} /></>;
      case 'orientation': return <><path d="M3 12l6-8 3 1" {...s} /><path d="M10 13.5h4" {...s} /></>;
      case 'currency': return <><rect x="1.5" y="4.5" width="13" height="7" rx="1" {...s} /><circle cx="8" cy="8" r="1.8" {...s} /></>;
      case 'percent': return <><circle cx="4.5" cy="4.5" r="1.8" {...s} /><circle cx="11.5" cy="11.5" r="1.8" {...s} /><path d="M12.5 3L3.5 13" {...s} /></>;
      case 'comma': return <text x="8" y="12" textAnchor="middle" fontSize="11" fontWeight="700" fill="currentColor" fontFamily="Selawik, sans-serif">,</text>;
      case 'decInc': return <text x="8" y="11" textAnchor="middle" fontSize="6.5" fill="currentColor" fontFamily="Selawik, sans-serif">←.00</text>;
      case 'decDec': return <text x="8" y="11" textAnchor="middle" fontSize="6.5" fill="currentColor" fontFamily="Selawik, sans-serif">.00→</text>;
      case 'condFormat': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><rect x="3.5" y="3.5" width="4" height="4" fill="#d13438" /><rect x="8.5" y="8.5" width="4" height="4" fill={G} /><path d="M2.5 8h11M8 2.5v11" {...s} /></>;
      case 'formatTable': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><rect x="2.5" y="2.5" width="11" height="3" fill="#5b9bd5" /><path d="M2.5 8.5h11M2.5 11h11M6.5 2.5v11M10 2.5v11" {...s} /></>;
      case 'cellStyles': return <><rect x="2.5" y="4.5" width="11" height="7" {...s} /><rect x="3.5" y="5.5" width="5" height="5" fill="#ffc7ce" /></>;
      case 'insertCells': return <><rect x="2.5" y="5.5" width="11" height="8" {...s} /><path d="M8 1.5v5M5.5 4L8 1.5 10.5 4" stroke={G} fill="none" /></>;
      case 'deleteCells': return <><rect x="2.5" y="5.5" width="11" height="8" {...s} /><path d="M5 1.5l5 4M10 1.5l-5 4" stroke="#d13438" fill="none" /></>;
      case 'format': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 6.5h11M6.5 2.5v11" {...s} /></>;
      case 'autosum': return <path d="M12 3.5H4l4.5 4.5L4 12.5h8" {...s} strokeWidth={1.3} />;
      case 'fillDown': return <><rect x="3.5" y="2.5" width="9" height="11" {...s} /><path d="M8 5v6M5.5 8.5L8 11l2.5-2.5" stroke={B} fill="none" /></>;
      case 'clear': return <><path d="M9 3l4 4-6 6H4l-1.5-1.5z" {...s} /><path d="M6 13h8" {...s} /></>;
      case 'sort': return <><path d="M4 2.5v11M2 11.5l2 2 2-2" {...s} /><text x="11" y="7" textAnchor="middle" fontSize="6" fill="currentColor" fontFamily="Selawik, sans-serif">A</text><text x="11" y="14" textAnchor="middle" fontSize="6" fill="currentColor" fontFamily="Selawik, sans-serif">Z</text></>;
      case 'sortZA': return <><path d="M4 2.5v11M2 11.5l2 2 2-2" {...s} /><text x="11" y="7" textAnchor="middle" fontSize="6" fill="currentColor" fontFamily="Selawik, sans-serif">Z</text><text x="11" y="14" textAnchor="middle" fontSize="6" fill="currentColor" fontFamily="Selawik, sans-serif">A</text></>;
      case 'filter': return <path d="M2 3h12l-4.5 5.5V13l-3 1.5v-6z" {...s} />;
      case 'find': return <><circle cx="6.5" cy="6.5" r="4" {...s} /><path d="M9.5 9.5l4 4" {...s} /></>;
      case 'chart': return <><path d="M2.5 13.5h11" {...s} /><rect x="3.5" y="8" width="2" height="5" fill={B} /><rect x="7" y="4" width="2" height="9" fill={G} /><rect x="10.5" y="6" width="2" height="7" fill="#ffb900" /></>;
      case 'pie': return <><circle cx="8" cy="8" r="5.5" {...s} /><path d="M8 2.5V8h5.5" fill="none" stroke="currentColor" /><path d="M8 8V2.5A5.5 5.5 0 0 1 13.5 8z" fill={G} /></>;
      case 'line': return <><path d="M2.5 13.5h11M2.5 2.5v11" {...s} /><path d="M3.5 11l3-4 3 2 4-5" stroke={B} fill="none" /></>;
      case 'pivot': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 6h11M6 2.5v11" {...s} /><path d="M8.5 9.5h3M10 8v3" stroke={G} fill="none" /></>;
      case 'table': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 6h11M2.5 9.5h11M6.5 2.5v11M10 2.5v11" {...s} /></>;
      case 'picture': return <><rect x="2.5" y="3.5" width="11" height="9" {...s} /><path d="M3 11.5l3.5-4 2.5 3 1.5-1.5 2.5 2.5" {...s} /><circle cx="10.5" cy="6" r="1" fill="currentColor" /></>;
      case 'shapes': return <><circle cx="5.5" cy="5.5" r="3" {...s} /><rect x="7.5" y="7.5" width="6" height="6" {...s} /></>;
      case 'link': return <><path d="M7 9a2.5 2.5 0 0 0 3.5 0l2-2a2.5 2.5 0 0 0-3.5-3.5l-.8.8" {...s} /><path d="M9 7a2.5 2.5 0 0 0-3.5 0l-2 2A2.5 2.5 0 0 0 7 12.5l.8-.8" {...s} /></>;
      case 'textbox': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M5 5.5h6M8 5.5v6" {...s} /></>;
      case 'symbol': return <text x="8" y="12.5" textAnchor="middle" fontSize="12" fill="currentColor" fontFamily="Selawik, sans-serif">Ω</text>;
      case 'function': return <text x="8" y="12" textAnchor="middle" fontSize="11" fontStyle="italic" fill="currentColor" fontFamily="Caladea, serif">fx</text>;
      case 'calc': return <><rect x="3.5" y="1.5" width="9" height="13" rx="1" {...s} /><rect x="5" y="3" width="6" height="3" {...s} /><path d="M5.5 8.5h1M9.5 8.5h1M5.5 11h1M9.5 11h1" {...s} /></>;
      case 'book': return <><path d="M2.5 3.5h4.5a1 1 0 0 1 1 1v9a1.5 1.5 0 0 0-1.5-1.5h-4z" {...s} /><path d="M13.5 3.5H9a1 1 0 0 0-1 1v9a1.5 1.5 0 0 1 1.5-1.5h4z" {...s} /></>;
      case 'logical': return <path d="M3 8.5l3 3 7-7" {...s} />;
      case 'text': return <text x="8" y="12" textAnchor="middle" fontSize="10" fill="currentColor" fontFamily="Selawik, sans-serif">A</text>;
      case 'clock': return <><circle cx="8" cy="8" r="5.5" {...s} /><path d="M8 4.5V8l2.5 1.5" {...s} /></>;
      case 'lookup': return <><circle cx="7" cy="7" r="3.5" {...s} /><path d="M9.5 9.5l4 4" {...s} /><path d="M5.5 7h3" {...s} /></>;
      case 'math': return <text x="8" y="12" textAnchor="middle" fontSize="10" fill="currentColor" fontFamily="Selawik, sans-serif">θ</text>;
      case 'more': return <><circle cx="4" cy="8" r=".9" fill="currentColor" /><circle cx="8" cy="8" r=".9" fill="currentColor" /><circle cx="12" cy="8" r=".9" fill="currentColor" /></>;
      case 'names': return <><rect x="2.5" y="4.5" width="11" height="7" {...s} /><path d="M4.5 8h5" {...s} /></>;
      case 'trace': return <><circle cx="4" cy="11" r="1.5" {...s} /><path d="M5 10l6-6M9 4h2v2" stroke={B} fill="none" /></>;
      case 'showFormulas': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><text x="8" y="11" textAnchor="middle" fontSize="7" fontStyle="italic" fill="currentColor" fontFamily="Caladea, serif">fx</text></>;
      case 'calcOptions': return <><rect x="3.5" y="1.5" width="9" height="13" rx="1" {...s} /><path d="M5.5 5h5M5.5 8h5M5.5 11h3" {...s} /></>;
      case 'getData': return <><ellipse cx="8" cy="4" rx="5" ry="2" {...s} /><path d="M3 4v8c0 1.1 2.2 2 5 2s5-.9 5-2V4" {...s} /></>;
      case 'refresh': return <><path d="M13 8a5 5 0 1 1-1.5-3.6" {...s} /><path d="M12 2v3h-3" {...s} /></>;
      case 'textToCols': return <><rect x="2.5" y="3.5" width="5" height="9" {...s} /><rect x="9.5" y="3.5" width="4" height="9" {...s} /></>;
      case 'dupes': return <><rect x="2.5" y="2.5" width="8" height="8" {...s} /><path d="M7 13.5l2-2 4.5-4.5" stroke="#d13438" fill="none" /></>;
      case 'validation': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M5 8l2 2 4-4" stroke={G} fill="none" /></>;
      // ---- title bar, formula bar, sheet bar and status bar glyphs (light theme)
      case 'saveQat': return <><path d="M3 2.5h8.2l1.8 1.8V13.5H3z" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinejoin="round" /><path d="M5.2 2.5v3.3h4.6V2.5M5 13.5V9.3h6v4.2" fill="none" stroke="currentColor" strokeWidth="1.1" strokeLinejoin="round" /></>;
      // Fluent ArrowUndo/ArrowRedo as the Office QAT draws them: an L-shaped head and a hooked tail
      case 'undoQat': return <><path d="M3.6 2v5.2h4.8" {...s} strokeWidth={1.1} /><path d="M4.4 6.4L7.6 3.2C8.8 2.1 11.4 2 12.6 3.8c1 1.5.7 3.5-.4 4.8L7.6 14.4" {...s} strokeWidth={1.1} /></>;
      case 'redoQat': return <><path d="M13.2 2v5.2H8.4" {...s} strokeWidth={1.1} /><path d="M12.4 6.4L9.2 3.2C8 2.1 5.4 2 4.2 3.8c-1 1.5-.7 3.5.4 4.8l4.6 5.8" {...s} strokeWidth={1.1} /></>;
      case 'fbarGrip': return <><rect x="6" y="3" width="2" height="2" fill="currentColor" /><rect x="6" y="7" width="2" height="2" fill="currentColor" /><rect x="6" y="11" width="2" height="2" fill="currentColor" /></>;
      case 'qatMore': return <><path d="M4.5 4.5h7" {...s} /><path d="M5 8l3 3 3-3" {...s} /></>;
      // Office's small drop-down chevron: 8 x 4 px at 45 degrees (drawn at size 12)
      case 'chevronV': return <path d="M2.67 5.33L8 10.67l5.33-5.33" {...s} strokeWidth={1.6} strokeLinecap="butt" />;
      case 'chevronDownThin': return <path d="M3.5 6l4.5 4.5L12.5 6" {...s} />;
      case 'searchQat': return <><circle cx="9.5" cy="6.5" r="4" {...s} /><path d="M6.6 9.4l-4.1 4.1" {...s} /></>;
      case 'commentPill': return <path d="M2.5 3.5h11v7.5H7.2l-3 2.6V11H2.5z" {...s} />;
      case 'sharePill': return <><path d="M9.5 3.5l4 3.6-4 3.6V8.8c-2.6 0-4.4.8-5.8 2.7.5-3 2.4-5 5.8-5.2z" {...s} /><path d="M7 3.5H3.5v9h9v-2" {...s} /></>;
      case 'vdots': return <><circle cx="8" cy="3.5" r="1.1" fill="currentColor" /><circle cx="8" cy="8" r="1.1" fill="currentColor" /><circle cx="8" cy="12.5" r="1.1" fill="currentColor" /></>;
      case 'fxCancel': return <path d="M3.5 3.5l9 9M12.5 3.5l-9 9" {...s} />;
      case 'fxEnter': return <path d="M2.5 8.5l3.5 3.5 7.5-8" {...s} />;
      case 'fx': return <text x="8" y="12.5" textAnchor="middle" fontSize="12.5" fontStyle="italic" fill="currentColor" fontFamily="Caladea, 'Liberation Serif', serif">fx</text>;
      // sheet-bar and status-bar glyphs, drawn at 16 px on whole pixels
      case 'addSheet': return <><rect x="7" y="3" width="1" height="11" fill="currentColor" /><rect x="2" y="8" width="11" height="1" fill="currentColor" /></>;
      case 'sheetPrev': return <path d="M9.5 4.5l-4 4 4 4" {...s} strokeWidth={1.3} />;
      case 'sheetNext': return <path d="M5.5 4.5l4 4-4 4" {...s} strokeWidth={1.3} />;
      case 'splitGrip': return <><rect x="7" y="5" width="2" height="2" fill="currentColor" /><rect x="7" y="9" width="2" height="2" fill="currentColor" /><rect x="7" y="13" width="2" height="2" fill="currentColor" /></>;
      case 'zoomOut': return <rect x="5" y="6" width="8" height="1" fill="currentColor" />;
      case 'zoomIn': return <><rect x="5" y="6" width="8" height="1" fill="currentColor" /><rect x="9" y="3" width="1" height="8" fill="currentColor" opacity={0.6} /></>;
      case 'navPrev': return <path d="M10 3.5L5.5 8l4.5 4.5" {...s} />;
      case 'navNext': return <path d="M6 3.5L10.5 8 6 12.5" {...s} />;
      case 'plusThin': return <path d="M8 2.5v11M2.5 8h11" {...s} />;
      // Fluent 'accessibility checkmark': figure with raised shoulders, check mark at the lower right
      case 'accessStatus': return <><circle cx="6.5" cy="2.8" r="1.5" {...s} /><path d="M1.5 4.6c1.6.7 3.3 1 5 1s3.4-.3 5-1M4.5 5.8v3.4L3.2 14M8.5 5.8v3" {...s} /><path d="M8.6 12.4l1.9 1.9 3.9-4.3" {...s} /></>;
      case 'normalViewSb': return <><rect x="2.5" y="1.5" width="12" height="12" {...s} /><path d="M2.5 5.5h12M2.5 9.5h12M6.5 1.5v12M10.5 1.5v12" {...s} /></>;
      case 'pageLayoutViewSb': return <><rect x="2.5" y="1.5" width="13" height="12" {...s} /><rect x="5.5" y="3.5" width="7" height="8" {...s} /><path d="M7 5.5h4M7 7.5h4M7 9.5h4" {...s} strokeLinecap="butt" /></>;
      case 'pageBreakViewSb': return <><rect x="2.5" y="1.5" width="13" height="12" {...s} /><path d="M6.5 1.5v7M10.5 1.5v7M2.5 8.5h8" {...s} strokeLinecap="butt" /></>;
      case 'pvInfo': return <><circle cx="8" cy="8" r="6" {...s} /><path d="M8 7.2v4.3M8 4.8v.4" {...s} /></>;
      case 'triUp': return <path d="M8 5.3l3.9 5.2H4.1z" fill="currentColor" stroke="currentColor" strokeLinejoin="round" />;
      case 'triDown': return <path d="M8 10.7L4.1 5.5h7.8z" fill="currentColor" stroke="currentColor" strokeLinejoin="round" />;
      case 'triLeft': return <path d="M5.3 8l5.2-3.9v7.8z" fill="currentColor" stroke="currentColor" strokeLinejoin="round" />;
      case 'triRight': return <path d="M10.7 8L5.5 11.9V4.1z" fill="currentColor" stroke="currentColor" strokeLinejoin="round" />;
      case 'quickAnalysis': return <><rect x="2.5" y="4.5" width="9" height="9" fill="#fff" stroke="#616161" /><path d="M4 7h6M4 9h6M4 11h4" stroke="#616161" fill="none" /><path d="M12 1.5l-2.2 3.4h1.7l-1.2 3 3-4h-1.8l1.2-2.4z" fill="#e8a33d" /></>;
      case 'copilotFab': return <><defs><linearGradient id="xlfab1" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stopColor="#3e8ef7" /><stop offset=".55" stopColor="#7b5cf0" /><stop offset="1" stopColor="#e85aa6" /></linearGradient><linearGradient id="xlfab2" x1="1" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#ffb547" /><stop offset=".6" stopColor="#f56f8f" /><stop offset="1" stopColor="#3fb8f0" /></linearGradient></defs><path d="M5.2 3.2h3.6c.8 0 1.4.5 1.6 1.2l2.4 8c.2.7-.3 1.4-1 1.4H8.2c-.8 0-1.4-.5-1.6-1.2l-2.4-8c-.2-.7.3-1.4 1-1.4z" fill="url(#xlfab1)" /><path d="M10.8 12.8H7.2c-.8 0-1.4-.5-1.6-1.2L3.2 3.6l5.9 7.8c.4.6 1.1 1 1.7 1.4z" fill="url(#xlfab2)" opacity=".85" /></>;
      case 'menuDeleteSheet': return <><path d="M1.5 2.5h10v8M1.5 2.5v10h7M1.5 5.5h10M1.5 8.5h7M5 2.5v10M8.5 2.5v6" fill="none" stroke="#424242" /><path d="M10 10l4.5 4.5M14.5 10L10 14.5" stroke="#d13438" strokeWidth="1.3" /></>;
      case 'menuRenameSheet': return <><path d="M1.5 2.5h12v4M1.5 2.5v10h6M1.5 5.5h12M1.5 8.5h6M5 2.5v10M9 2.5v3" fill="none" stroke="#424242" /><path d="M8.5 14.5l.5-2.3 4.6-4.6 1.8 1.8-4.6 4.6z" fill="none" stroke="#2b7cd3" /></>;
      case 'menuViewCode': return <><rect x="1.5" y="2.5" width="13" height="11" fill="none" stroke="#424242" /><path d="M3.5 6h2M3.5 8.5h2M3.5 11h2" stroke="#424242" /><circle cx="10" cy="8.5" r="2.3" fill="none" stroke="#2b7cd3" /><path d="M11.7 10.2l2 2" stroke="#2b7cd3" /></>;
      case 'menuProtectSheet': return <><path d="M1.5 2.5h12v4M1.5 2.5v10h6M1.5 5.5h12M1.5 8.5h6M5 2.5v10M9 2.5v3" fill="none" stroke="#424242" /><rect x="9" y="10" width="6" height="4.5" fill="#e8892c" /><path d="M10.3 10V8.8a1.7 1.7 0 0 1 3.4 0V10" fill="none" stroke="#e8892c" /></>;
      case 'menuLinkSheet': return <><path d="M1.5 2.5h12v5M1.5 2.5v10h5M1.5 5.5h12M1.5 8.5h5M5 2.5v10M9 2.5v3" fill="none" stroke="#424242" /><path d="M9.5 11.5h-.8a1.6 1.6 0 0 1 0-3.2h1.8a1.6 1.6 0 0 1 1.5 1M12.5 10.5h.8a1.6 1.6 0 0 1 0 3.2h-1.8a1.6 1.6 0 0 1-1.5-1" fill="none" stroke="#2b7cd3" /></>;
      case 'whatif': return <text x="8" y="12" textAnchor="middle" fontSize="11" fill="currentColor" fontFamily="Selawik, sans-serif">?</text>;
      case 'group': return <><path d="M3 2.5v11M3 2.5h2M3 13.5h2" {...s} /><path d="M7 5h6M7 8h6M7 11h6" {...s} /></>;
      case 'spelling': return <><text x="6" y="8" textAnchor="middle" fontSize="6.5" fill="currentColor" fontFamily="Selawik, sans-serif">ABC</text><path d="M6 11.5l2 2 5-5" stroke={G} fill="none" /></>;
      case 'thesaurus': return <path d="M3 2.5h8l2 2v9H3z M5.5 6h5M5.5 8.5h5M5.5 11h3" {...s} />;
      case 'access': return <><circle cx="8" cy="3.5" r="1.5" {...s} /><path d="M3 6h10M8 6v4M8 10l-2.5 4M8 10l2.5 4" {...s} /></>;
      case 'comment': return <path d="M2.5 3.5h11v7.5h-6l-3 2.5V11H2.5z" {...s} />;
      case 'note': return <><path d="M2.5 2.5h11v8l-3 3h-8z" {...s} /><path d="M10.5 13.5v-3h3" {...s} /></>;
      case 'protect': return <><rect x="3.5" y="7" width="9" height="7" rx="1" {...s} /><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" {...s} /></>;
      case 'normalView': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 6h11M2.5 9.5h11M6.5 2.5v11M10 2.5v11" {...s} /></>;
      case 'pageLayoutView': return <><rect x="3.5" y="1.5" width="9" height="13" {...s} /><path d="M5.5 4.5h5M5.5 7h5M5.5 9.5h5" {...s} /></>;
      case 'pageBreakView': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M8 2.5v11M2.5 8h11" stroke={B} strokeDasharray="1.5 1" fill="none" /></>;
      case 'freeze': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 6h11M6 2.5v11" stroke={B} strokeWidth="1.5" fill="none" /></>;
      case 'zoom': return <><circle cx="6.5" cy="6.5" r="4" {...s} /><path d="M9.5 9.5l4 4M4.5 6.5h4M6.5 4.5v4" {...s} /></>;
      case 'zoom100': return <text x="8" y="11" textAnchor="middle" fontSize="6.5" fill="currentColor" fontFamily="Selawik, sans-serif">100</text>;
      case 'zoomSel': return <><rect x="2.5" y="2.5" width="11" height="11" strokeDasharray="1.5 1" {...s} /><circle cx="8" cy="8" r="2.5" {...s} /></>;
      case 'newWindow': return <><rect x="2.5" y="4.5" width="9" height="9" {...s} /><path d="M5.5 4.5v-2h8v8h-2" {...s} /></>;
      case 'arrange': return <><rect x="2.5" y="2.5" width="5" height="11" {...s} /><rect x="8.5" y="2.5" width="5" height="11" {...s} /></>;
      case 'switch': return <><rect x="2.5" y="4.5" width="8" height="8" {...s} /><rect x="5.5" y="2.5" width="8" height="8" {...s} /></>;
      case 'macros': return <><rect x="2.5" y="3.5" width="11" height="9" {...s} /><path d="M6 6.5l3 1.5-3 1.5z" fill="currentColor" /></>;
      case 'help': return <><circle cx="8" cy="8" r="5.5" {...s} /><path d="M6.5 6.5a1.5 1.5 0 1 1 2 1.4c-.4.2-.5.5-.5.9v.5M8 11v.3" {...s} /></>;
      case 'feedback': return <><path d="M2.5 3.5h11v7.5h-6l-3 2.5V11H2.5z" {...s} /><path d="M8 5.5v2.5M8 9.5v.3" {...s} /></>;
      case 'training': return <><rect x="2.5" y="3.5" width="11" height="8" {...s} /><path d="M6.5 6l3 1.5-3 1.5z" fill="currentColor" /><path d="M5.5 13.5h5" {...s} /></>;
      case 'themes': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><rect x="3.5" y="9" width="2.5" height="3.5" fill="#4472c4" /><rect x="6.5" y="9" width="2.5" height="3.5" fill="#ed7d31" /><rect x="9.5" y="9" width="2.5" height="3.5" fill="#a5a5a5" /><path d="M4 5h8" {...s} /></>;
      case 'margins': return <><rect x="3.5" y="1.5" width="9" height="13" {...s} /><path d="M5.5 1.5v13M10.5 1.5v13M3.5 4h9M3.5 12h9" stroke={B} strokeDasharray="1 1" fill="none" /></>;
      case 'orientationPage': return <><rect x="2.5" y="4.5" width="11" height="8" {...s} /><path d="M5 1.5v3" {...s} /></>;
      case 'size': return <><rect x="4.5" y="1.5" width="8" height="12" {...s} /><path d="M2.5 3.5v11h8" {...s} /></>;
      case 'printArea': return <><rect x="2.5" y="2.5" width="11" height="11" strokeDasharray="2 1" {...s} /><path d="M5 6h6M5 9h4" {...s} /></>;
      case 'breaks': return <><rect x="3.5" y="1.5" width="9" height="5" {...s} /><rect x="3.5" y="9.5" width="9" height="5" {...s} /><path d="M1.5 8h13" stroke={B} strokeDasharray="1.5 1" fill="none" /></>;
      case 'background': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 11l4-4 3 3 4-4" {...s} /></>;
      case 'printTitles': return <><rect x="2.5" y="2.5" width="11" height="11" {...s} /><path d="M2.5 5.5h11" stroke={B} strokeWidth="2" fill="none" /></>;
      case 'share': return <><path d="M9 3.5l4 3.5-4 3.5V8.5C5.5 8.5 4 10 3 12.5 3 9 4.5 5.5 9 5.5z" {...s} /></>;
      case 'plus': return <path d="M8 3v10M3 8h10" {...s} />;
      case 'minus': return <path d="M3 8h10" {...s} />;
      case 'navFirst': return <path d="M5 4v8M11 4L7 8l4 4" {...s} />;
      case 'navLast': return <path d="M11 4v8M5 4l4 4-4 4" {...s} />;
      case 'info': return <><circle cx="8" cy="8" r="6" {...s} /><path d="M8 7v4.5M8 4.8v.4" {...s} /></>;
      case 'folder': return <path d="M1.5 4.5v8.5h13V5.5H7.5L6 3.5H1.5z" {...s} />;
      case 'home': return <><path d="M2.5 7.5L8 2.5l5.5 5v6h-11z" {...s} /><path d="M6.5 13.5v-4h3v4" {...s} /></>;
      case 'new': return <path d="M3.5 1.5h6l3 3v10h-9z M9.5 1.5v3h3" {...s} />;
      case 'open': return <path d="M1.5 4v9h11l2-6.5H4L1.5 13M1.5 4V2.5h4.5l1.5 1.5h5v2.5" {...s} />;
      case 'print': return <><rect x="4.5" y="1.5" width="7" height="4" {...s} /><rect x="1.5" y="5.5" width="13" height="6" rx="1" {...s} /><rect x="4.5" y="9.5" width="7" height="5" fill="#fff" stroke="currentColor" /></>;
      case 'export': return <><path d="M3.5 1.5h6l3 3v10h-9z" {...s} /><path d="M8 6.5v6M5.5 9l2.5-2.5L10.5 9" {...s} /></>;
      case 'back': return <><circle cx="8" cy="8" r="6.5" {...s} /><path d="M9.5 5L6.5 8l3 3" {...s} /></>;
      case 'goTo': return <path d="M2.5 8h9M8.5 4.5L12 8l-3.5 3.5" {...s} />;
      case 'dropdown': return <path d="M4.5 6.5l3.5 3.5 3.5-3.5" fill="currentColor" stroke="none" />;
      case 'copilotless': return null;
      default: return <rect x="3" y="3" width="10" height="10" {...s} />;
    }
  })();
  return <svg className="xl-icon" width={size} height={size} viewBox="0 0 16 16" aria-hidden style={{ display: 'block', flex: 'none', ...style }}>{body}</svg>;
}
