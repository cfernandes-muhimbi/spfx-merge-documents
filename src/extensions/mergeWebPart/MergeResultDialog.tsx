import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';

interface IResultContentProps {
  mergedFileUrl: string;
  onOpen: () => void;
  onClose: () => void;
}

const MergeResultContent: React.FC<IResultContentProps> = ({ mergedFileUrl, onOpen, onClose }) => (
  <div style={{ padding: '20px', minWidth: '480px', maxWidth: '680px' }}>
    <p style={{ fontSize: '14px', color: '#323130', marginBottom: '8px' }}>
      Documents bundled successfully. Your merged document is ready:
    </p>
    <div
      style={{
        padding: '10px 12px',
        background: '#f3f2f1',
        border: '1px solid #edebe9',
        borderRadius: '2px',
        marginBottom: '20px',
        wordBreak: 'break-all',
        fontSize: '13px',
        color: '#0078d4'
      }}
    >
      {mergedFileUrl}
    </div>
    <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
      <PrimaryButton text="Open Document" onClick={onOpen} />
      <DefaultButton text="Close" onClick={onClose} />
    </div>
  </div>
);

export class MergeResultDialog extends BaseDialog {
  public mergedFileUrl: string = '';

  public render(): void {
    ReactDOM.render(
      <MergeResultContent
        mergedFileUrl={this.mergedFileUrl}
        onOpen={this._handleOpen}
        onClose={this._handleClose}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  private _handleOpen = async (): Promise<void> => {
    window.open(this.mergedFileUrl, '_blank');
    await this.close();
  };

  private _handleClose = async (): Promise<void> => {
    await this.close();
  };
}
