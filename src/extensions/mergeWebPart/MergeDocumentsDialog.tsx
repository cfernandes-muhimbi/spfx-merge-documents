import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';

export interface IDocumentItem {
  name: string;
  serverRelativeUrl: string;
}

interface IDialogContentProps {
  documents: IDocumentItem[];
  onBundle: (orderedUrls: string[]) => void;
  onCancel: () => void;
}

interface IDialogContentState {
  documents: IDocumentItem[];
  dragOverIndex: number | null;
}

class MergeDialogContent extends React.Component<IDialogContentProps, IDialogContentState> {
  private _dragIndex: number | null = null;

  constructor(props: IDialogContentProps) {
    super(props);
    this.state = {
      documents: [...props.documents],
      dragOverIndex: null
    };
  }

  private _onDragStart = (e: React.DragEvent<HTMLDivElement>, index: number): void => {
    this._dragIndex = index;
    e.dataTransfer.effectAllowed = 'move';
  };

  private _onDragOver = (e: React.DragEvent<HTMLDivElement>, index: number): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    this.setState({ dragOverIndex: index });
  };

  private _onDrop = (e: React.DragEvent<HTMLDivElement>, dropIndex: number): void => {
    e.preventDefault();
    if (this._dragIndex === null || this._dragIndex === dropIndex) {
      this.setState({ dragOverIndex: null });
      return;
    }
    const newDocs = [...this.state.documents];
    const dragged = newDocs.splice(this._dragIndex, 1)[0];
    newDocs.splice(dropIndex, 0, dragged);
    this._dragIndex = null;
    this.setState({ documents: newDocs, dragOverIndex: null });
  };

  private _onDragEnd = (): void => {
    this._dragIndex = null;
    this.setState({ dragOverIndex: null });
  };

  public render(): React.ReactElement {
    const { documents, dragOverIndex } = this.state;

    return (
      <div style={{ padding: '20px', minWidth: '520px', maxWidth: '700px' }}>
        <p style={{ color: '#605e5c', marginBottom: '8px', fontSize: '14px' }}>
          Drag items to change the order
        </p>
        <div style={{ border: '1px solid #edebe9', marginBottom: '20px' }}>
          {documents.map((doc, index) => (
            <div
              key={doc.serverRelativeUrl}
              draggable
              onDragStart={(e) => this._onDragStart(e, index)}
              onDragOver={(e) => this._onDragOver(e, index)}
              onDrop={(e) => this._onDrop(e, index)}
              onDragEnd={this._onDragEnd}
              style={{
                padding: '10px 16px',
                borderBottom: index < documents.length - 1 ? '1px solid #edebe9' : 'none',
                cursor: 'grab',
                background: dragOverIndex === index ? '#f3f2f1' : '#ffffff',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                userSelect: 'none'
              }}
            >
              <span style={{ color: '#a19f9d', fontSize: '18px', lineHeight: '1' }}>&#8801;</span>
              <span style={{ fontSize: '14px', color: '#323130' }}>{doc.name}</span>
            </div>
          ))}
        </div>
        <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
          <PrimaryButton
            text="Bundle"
            onClick={() => this.props.onBundle(this.state.documents.map(d => d.serverRelativeUrl))}
          />
          <DefaultButton text="Cancel" onClick={this.props.onCancel} />
        </div>
      </div>
    );
  }
}

export class MergeDocumentsDialog extends BaseDialog {
  public documents: IDocumentItem[] = [];
  public orderedUrls: string[] = [];
  public wasBundled: boolean = false;

  public render(): void {
    ReactDOM.render(
      <MergeDialogContent
        documents={this.documents}
        onBundle={this._handleBundle}
        onCancel={this._handleCancel}
      />,
      this.domElement
    );
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  protected onAfterClose(): void {
    super.onAfterClose();
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  private _handleBundle = async (orderedUrls: string[]): Promise<void> => {
    this.orderedUrls = orderedUrls;
    this.wasBundled = true;
    await this.close();
  };

  private _handleCancel = async (): Promise<void> => {
    this.wasBundled = false;
    await this.close();
  };
}
