import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { Dialog } from '@microsoft/sp-dialog';
import { MergeDocumentsDialog, type IDocumentItem } from './MergeDocumentsDialog';
import { MergeResultDialog } from './MergeResultDialog';
import { MergeLoadingDialog } from './MergeLoadingDialog';

export interface IMergeWebPartCommandSetProperties {
  // No custom properties needed
}

const LOG_SOURCE: string = 'MergeWebPartCommandSet';

// ─── Update these two constants to match your environment ────────────────────

// Power Automate HTTP trigger URL — replace with your flow's HTTP POST URL
const FLOW_URL: string = 'https://<YOUR-FLOW-TRIGGER-URL>';

// Full URL of the Document Library where merged documents are saved
// e.g. https://<tenant>.sharepoint.com/sites/<site>/Shared%20Documents
const MERGED_DOCS_LIBRARY_URL: string = 'https://<YOUR-TENANT>.sharepoint.com/sites/<YOUR-SITE>/<YOUR-LIBRARY>';

// ─────────────────────────────────────────────────────────────────────────────

export default class MergeWebPartCommandSet extends BaseListViewCommandSet<IMergeWebPartCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MergeWebPartCommandSet');

    const mergeCommand: Command = this.tryGetCommand('COMMAND_1');
    mergeCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this._showMergeDialog().catch(err => {
          console.error('MergeWebPart error:', err);
          Dialog.alert(`Error opening Merge dialog: ${err?.message ?? err}`).catch(() => { /* ignore */ });
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async _showMergeDialog(): Promise<void> {
    const selectedRows = this.context.listView.selectedRows;

    console.log('MergeWebPart: selected rows count:', selectedRows?.length);

    if (!selectedRows || selectedRows.length < 2) {
      await Dialog.alert('Please select 2 or more documents.');
      return;
    }

    const documents: IDocumentItem[] = selectedRows.map(row => {
      const name = (row.getValueByName('FileLeafRef') ?? row.getValueByName('Title') ?? 'Unknown') as string;
      const serverRelativeUrl = (row.getValueByName('FileRef') ?? '') as string;
      console.log('MergeWebPart: document:', name, serverRelativeUrl);
      return { name, serverRelativeUrl };
    });

    const dialog = new MergeDocumentsDialog();
    dialog.documents = documents;

    await dialog.show();

    if (dialog.wasBundled) {
      await this._triggerFlow(dialog.orderedUrls);
    }
  }

  private async _triggerFlow(serverRelativeUrls: string[]): Promise<void> {
    const loadingDialog = new MergeLoadingDialog();
    // Show without await — isBlocking:true keeps it open until we close() it
    const loadingPromise = loadingDialog.show();

    let mergedFileUrl: string | undefined;
    let errorMessage: string | undefined;

    try {
      console.log('MergeWebPart: triggering flow with URLs:', serverRelativeUrls);

      const client = await this.context.aadHttpClientFactory
        .getClient('https://service.flow.microsoft.com/');

      const response = await client.post(FLOW_URL, AadHttpClient.configurations.v1, {
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify({ fileUrls: serverRelativeUrls })
      });

      if (response.ok) {
        mergedFileUrl = await this._parseMergedFileUrl(response);
        Log.info(LOG_SOURCE, `Flow succeeded. Merged file: ${mergedFileUrl}`);
      } else {
        const text = await response.text();
        Log.warn(LOG_SOURCE, `Flow returned status: ${response.status}`);
        errorMessage = `Bundling failed (status ${response.status}):\n${text}`;
      }
    } catch (error) {
      console.error('MergeWebPart: flow error:', error);
      errorMessage = `Error triggering flow: ${(error as Error)?.message ?? error}`;
    } finally {
      await loadingDialog.close();
      await loadingPromise;
    }

    if (mergedFileUrl) {
      const resultDialog = new MergeResultDialog();
      resultDialog.mergedFileUrl = mergedFileUrl;
      await resultDialog.show();
    } else if (errorMessage) {
      await Dialog.alert(errorMessage);
    }
  }

  // Parses the flow response to extract the merged document URL.
  // The flow should return a JSON body with one of: mergedFileUrl, fileUrl, fileName, or name.
  // If the value is already a full URL it is used as-is; if it is a relative path or filename
  // it is appended to MERGED_DOCS_LIBRARY_URL.
  private async _parseMergedFileUrl(response: HttpClientResponse): Promise<string> {
    try {
      const json = await response.json() as Record<string, string>;
      console.log('MergeWebPart: flow response:', json);

      // Try common property names returned by Power Automate "Respond to a PowerApp or flow"
      const raw: string | undefined =
        json.mergedFileUrl ??
        json.fileUrl ??
        json.fileName ??
        json.name ??
        json.url;

      if (!raw) {
        // Flow returned no usable URL — fall back to the library root
        return MERGED_DOCS_LIBRARY_URL;
      }

      // If it's already an absolute URL, use it directly
      if (raw.startsWith('http://') || raw.startsWith('https://')) {
        return raw;
      }

      // Otherwise treat it as a filename / server-relative path and build the full URL
      const base = MERGED_DOCS_LIBRARY_URL.replace(/\/$/, '');
      const path = raw.startsWith('/') ? raw : `/${raw}`;
      return `${base}${path}`;
    } catch {
      // Response wasn't JSON — return the library URL as a fallback
      return MERGED_DOCS_LIBRARY_URL;
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const mergeCommand: Command = this.tryGetCommand('COMMAND_1');
    if (mergeCommand) {
      mergeCommand.visible = (this.context.listView.selectedRows?.length ?? 0) >= 2;
    }

    this.raiseOnChange();
  };
}
