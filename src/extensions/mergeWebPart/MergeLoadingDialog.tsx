import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

const LoadingContent: React.FC = () => (
  <div style={{ padding: '32px 40px', textAlign: 'center', minWidth: '300px' }}>
    <Spinner size={SpinnerSize.large} label="Bundling documents, please wait…" labelPosition="bottom" />
  </div>
);

export class MergeLoadingDialog extends BaseDialog {

  public render(): void {
    ReactDOM.render(<LoadingContent />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: true };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }
}
