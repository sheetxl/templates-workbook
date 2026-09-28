/**
 * Renders one template for its thumbnail: `?src=<category>/<name>.xlsx|.sxl|.csv&scheme=light|dark`.
 *
 * The grid is the same renderer the gallery opens the template in, so the thumbnail shows what the
 * user will get. It is rendered chromeless, with the row/column headers hidden, and
 * `config/thumbnails.ts` photographs only its cell area. `scheme=dark` is the dark app with the
 * dark grid turned on.
 *
 * `window.__thumbnail` reports the outcome to `config/thumbnails.ts`: `'ready'` once the grid has
 * painted, or an error message.
 */
import { createRoot } from 'react-dom/client';

import { WorkbookIO } from '@sheetxl/sdk';
import type { IWorkbook } from '@sheetxl/sdk';
import { IOPlugin } from '@sheetxl/io';
import { WorkbookElement } from '@sheetxl/studio-mui';

declare global {
  interface Window { __thumbnail?: 'ready' | { error: string } }
}

const params = new URLSearchParams(window.location.search);
const src = params.get('src');
const scheme = params.get('scheme') === 'dark' ? 'dark' : 'light';

/** The selection is parked here, far outside the first screen, so it is not in the picture. */
const PARKED_SELECTION = 'XFD1048576';
/** Zoomed out a little, so a thumbnail shows more of the template. */
const ZOOM = 70;

const reportError = (error: unknown): void => {
  window.__thumbnail = { error: error instanceof Error ? error.message : String(error) };
};

/**
 * Reads the template and sets up its view before the grid mounts. Moving the selection once the
 * grid is up scrolls the view after it (the grid follows the active cell); a selection that is
 * already there at mount leaves the template's saved scroll alone.
 */
async function readTemplate(url: string): Promise<IWorkbook> {
  await WorkbookIO.install(IOPlugin);
  const workbook = await WorkbookIO.read({ source: url });
  if (!workbook) throw new Error('the template did not read');
  const sheet = workbook.getSelectedSheet();
  sheet.getView().setZoomScale(ZOOM);
  await sheet.getRange(PARKED_SELECTION).select();
  return workbook;
}

const workbook = src ? readTemplate(new URL(src, window.location.href).href) : null;
workbook?.catch(reportError);

const App = () => {
  if (!workbook) {
    window.__thumbnail = { error: 'no ?src given' };
    return null;
  }
  return (
    <WorkbookElement
      workbook={workbook}
      showTabs={false}
      showHorizontalScrollbar={false}
      showVerticalScrollbar={false}
      colorScheme={scheme}
      gridColorScheme={scheme}
      propsSheet={{ showColumnHeaders: false, showRowHeaders: false }}
      style={{ position: 'absolute', inset: 0, colorScheme: scheme }}
      onWorkbookLoad={() => {
        // The load event fires one frame before the grid paints.
        requestAnimationFrame(() => requestAnimationFrame(() => { window.__thumbnail = 'ready'; }));
      }}
      renderWorkbookError={({ error }) => {
        reportError(error);
        return <></>;
      }}
    />
  );
};

createRoot(document.getElementById('root')!).render(<App/>);
