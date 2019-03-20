import * as React from "react";
import { DefaultButton } from "office-ui-fabric-react";
import Dropzone, { DropzoneProps } from "react-dropzone";

import { IShrelloItem } from "../../../models";
import styles from "../ShrelloForm.module.scss";

export interface IShrelloAttachmentsProps extends DropzoneProps {
  item: IShrelloItem;
  handleDrop: (files: File[]) => void;
}

export default class ShrelloAttachments extends React.Component<IShrelloAttachmentsProps, {
  files: File[],
  dropzoneActive: boolean
}> {
  private _dropzone: Dropzone;
  constructor(props: IShrelloAttachmentsProps) {
    super(props);
    this.state = {
      files: [],
      dropzoneActive: false
    };
    this._onDrop = this._onDrop.bind(this);
    this._onDragEnter = this._onDragEnter.bind(this);
    this._onDragLeave = this._onDragLeave.bind(this);
    this._onClick = this._onClick.bind(this);
  }

  private _onDragEnter = () => this.setState({ dropzoneActive: true });

  private _onDragLeave = () => this.setState({ dropzoneActive: false });

  private _onDrop = (files: File[]) => {
    // const comments: string = `Added attachments.`;
    this.props.handleDrop(files);
    this.setState({
      files,
      dropzoneActive: false
    });
  }

  private _onClick = (ev) => this._dropzone.open();

  public render(): JSX.Element {
    const {item} = this.props;
    const { files, dropzoneActive } = this.state;
    const overlayStyle: React.CSSProperties = {
      position: "absolute",
      top: 0,
      right: 0,
      bottom: 0,
      left: 0,
      padding: "2.5em 0",
      background: "rgba(0,0,0,0.5)",
      textAlign: "center",
      color: "#fff"
    };
    return (
      <div className="dropzone">
        <Dropzone
          disableClick
          ref={node => this._dropzone = node}
          style={{position: "relative", minHeight: "100px", paddingTop: "12px"}}
          onDrop={this._onDrop}
          onDragEnter={this._onDragEnter}
          onDragLeave={this._onDragLeave}
        >
          { dropzoneActive &&
            <div style={overlayStyle}>Drop files...</div> }
          <div>
            <p>Drag files here to add attachments...</p>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12 ms-lg6">
                <h2>Attached Files:</h2>
              { item && item.AttachmentFiles && item.AttachmentFiles.length &&
                <div className={styles.fileList}>
                  {
                    item.AttachmentFiles.map(file => {
                      return (
                        <div className={styles.file} key={file.FileName}>
                          <a href={file.ServerRelativeUrl} target="_blank">{file.FileName}</a>
                        </div>
                      );
                    })
                  }
                </div>
              }
              </div>
              <div className="ms-Grid-col ms-sm12 ms-lg6">
                <h2>Dropped files:</h2>
                <div className={styles.fileList}>
                  {
                    files.map(f => <div className={styles.file} key={f.name}>{f.name} - {f.size} bytes</div>)
                  }
                </div>
              </div>
            </div>
          </div>
        </Dropzone>
        <DefaultButton
          onClick={this._onClick}
          text="Open File Explorer"
        />
      </div>
    );
  }
}
