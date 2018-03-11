import {
  JupyterLab, JupyterLabPlugin
} from '@jupyterlab/application';

import '../style/index.css';


/**
 * Initialization data for the jupyterlab_xls extension.
 */
const extension: JupyterLabPlugin<void> = {
  id: 'jupyterlab_xls',
  autoStart: true,
  activate: (app: JupyterLab) => {
    console.log('JupyterLab extension jupyterlab_xls is activated!');
  }
};

export default extension;
