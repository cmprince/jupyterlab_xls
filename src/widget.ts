// Copyright (c) Jupyter Development Team.
// Distributed under the terms of the Modified BSD License.

// import * as dsv from 'd3-dsv';

import {
  ActivityMonitor, PathExt
} from '@jupyterlab/coreutils';

import {
  ABCWidgetFactory, DocumentRegistry
} from '@jupyterlab/docregistry';

import {
  PromiseDelegate
} from '@phosphor/coreutils';

import {
  DataGrid, JSONModel
} from '@phosphor/datagrid';

import {
  Message
} from '@phosphor/messaging';

import {
  PanelLayout, Widget
} from '@phosphor/widgets';

//import {
//  XLSToolbar
//} from './toolbar';

import * as XLSX from 'xlsx';

/**
 * The class name added to a XLS viewer.
 */
const XLS_CLASS = 'jp-XLSViewer';

/**
 * The class name added to a XLS viewer toolbar.
 */
//const XLS_VIEWER_CLASS = 'jp-XLSViewer-toolbar';

/**
 * The class name added to a XLS viewer datagrid.
 */
const XLS_GRID_CLASS = 'jp-XLSViewer-grid';

/**
 * The timeout to wait for change activity to have ceased before rendering.
 */
const RENDER_TIMEOUT = 1000;


/**
 * A viewer for XLS tables.
 */
export
class XLSViewer extends Widget implements DocumentRegistry.IReadyWidget {
  /**
   * Construct a new XLS viewer.
   */
  constructor(options: XLSViewer.IOptions) {
    super();

    let context = this._context = options.context;
    let layout = this.layout = new PanelLayout();

    this.addClass(XLS_CLASS);

    this._grid = new DataGrid();
    this._grid.addClass(XLS_GRID_CLASS);
    this._grid.headerVisibility = 'column';

      //    this._toolbar = new XLSToolbar({ selected: this._delimiter });
      //    this._toolbar.delimiterChanged.connect(this._onDelimiterChanged, this);
      //    this._toolbar.addClass(XLS_VIEWER_CLASS);
      //    layout.addWidget(this._toolbar);
    layout.addWidget(this._grid);

    context.pathChanged.connect(this._onPathChanged, this);
    this._onPathChanged();

    this._context.ready.then(() => {
      this._updateGrid();
      this._ready.resolve(undefined);
      // Throttle the rendering rate of the widget.
      this._monitor = new ActivityMonitor({
        signal: context.model.contentChanged,
        timeout: RENDER_TIMEOUT
      });
      this._monitor.activityStopped.connect(this._updateGrid, this);
    });
  }

  /**
   * The XLS widget's context.
   */
  get context(): DocumentRegistry.Context {
    return this._context;
  }

  /**
   * A promise that resolves when the XLS viewer is ready.
   */
  get ready() {
    return this._ready.promise;
  }

  /**
   * Dispose of the resources used by the widget.
   */
  dispose(): void {
    if (this._monitor) {
      this._monitor.dispose();
    }
    super.dispose();
  }

  /**
   * Handle `'activate-request'` messages.
   */
  protected onActivateRequest(msg: Message): void {
    this.node.tabIndex = -1;
    this.node.focus();
  }

  /**
   * Handle a change in delimiter.
   */
    //  private _onDelimiterChanged(sender: XLSToolbar, delimiter: string): void {
    //    this._delimiter = delimiter;
    //    this._updateGrid();
    //  }

  /**
   * Handle a change in path.
   */
  private _onPathChanged(): void {
    this.title.label = PathExt.basename(this._context.localPath);
  }

  /**
   * Create the json model for the grid.
   */
  private _updateGrid(): void {
    let text = this._context.model.toString();
    let [columns, data] = Private.parse(text, this._delimiter);
    let fields = columns.map(name => ({ name, type: 'string' }));
      //    this._grid.model = new JSONModel({ data, schema: { fields } });
  }

  private _context: DocumentRegistry.Context;
  private _grid: DataGrid;
    //  private _toolbar: XLSToolbar;
  private _monitor: ActivityMonitor<any, any> | null = null;
  private _delimiter = ',';
  private _ready = new PromiseDelegate<void>();
}


/**
 * A namespace for `XLSViewer` statics.
 */
export
namespace XLSViewer {
  /**
   * Instantiation options for XLS widgets.
   */
  export
  interface IOptions {
    /**
     * The document context for the XLS being rendered by the widget.
     */
    context: DocumentRegistry.Context;
  }
}


/**
 * A widget factory for XLS widgets.
 */
export
class XLSViewerFactory extends ABCWidgetFactory<XLSViewer, DocumentRegistry.IModel> {
  /**
   * Create a new widget given a context.
   */
  protected createNewWidget(context: DocumentRegistry.Context): XLSViewer {
    return new XLSViewer({ context });
  }
}


/**
 * The namespace for the module implementation details.
 */
namespace Private {
  /**
   * Parse DSV text with the given delimiter.
   *
   * @param text - The DSV text to parse.
   *
   * @param delimiter - The delimiter for parsing.
   *
   * @returns A tuple of `[columnNames, dataRows]`
   */
  export
  function parse(fname: string, delimiter: string): [string[], string[]] { // dsv.DSVRowString[]] {
    let columns: string[] = [];
      // let rowFn: RowFn | null = null;
    const wb: XLSX.WorkBook = readFile('test.xlsx');
    const ws: XLSX.WorkSheet = wb.Sheets[wb.SheetNames[0]];
      let rows = XLSX.utils.sheet_to_json(ws);
      //let columns = rows;
      //    let rows = dsv.dsvFormat(delimiter).parseRows(text, row => {
      //      if (rowFn) {
      //        return rowFn(row);
      //      }
      //      columns = uniquifyColumns(row);
      //      rowFn = makeRowFn(columns);
      //    });
    return [columns, rows];
  }

  /**
   * Replace duplicate column names with unique substitutes.
   */
  function uniquifyColumns(columns: string[]): string[] {
    let unique: string[] = [];
    let set: { [key: string]: boolean } = Object.create(null);
    for (let name of columns) {
      let uniqueName = name;
      for (let i = 1; uniqueName in set; ++i) {
        uniqueName = `${name}.${i}`;
      }
      set[uniqueName] = true;
      unique.push(uniqueName);
    }
    return unique;
  }

  /**
   * A type alias for a row conversion function.
   */
    //  type RowFn = (r: string[]) => dsv.DSVRowString;

  /**
   * Create a row conversion function for the given column names.
   */
  function makeRowFn(columns: string[]): RowFn {
    let pairs = columns.map((name, i) => `'${name}':r[${i}]`).join(',');
    return (new Function('r', `return {${pairs}};`)) as RowFn;
  }
}
