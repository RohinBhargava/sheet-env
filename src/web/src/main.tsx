import { createRoot } from 'react-dom/client';
import { App } from './App';
import { MenuHost, TooltipHost } from './shell/Menu';
import { installClipboard } from './shell/clipboard';
import './styles/fonts.css';
import './styles/menu.css';
import './styles/app.css';

installClipboard();

// No StrictMode: the workbook view's effects (load, event log) are written to run once per mount.
createRoot(document.getElementById('root')!).render(
  <>
    <App />
    <MenuHost />
    <TooltipHost />
  </>,
);
