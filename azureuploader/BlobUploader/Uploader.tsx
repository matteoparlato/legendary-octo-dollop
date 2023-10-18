import * as React from 'react';
import { DefaultButton, Label, ProgressIndicator } from '@fluentui/react';
import { IControlEvent } from './IControlEvent';
import SharePointUploader from './SharePointHelper';
import { validateDrive } from './drive-utils';

export interface IUploaderProps {
  storageName: string | null;
  storagePath: string | null;
  sasToken: string | null;
  label: string | null;
  uploadLabel: string | null;
  onEvent: (event: IControlEvent) => void;
}

type UploaderState = {
  uploading: Boolean,
  selectedFile: File | null,
  inputKey: React.Key,
  percentCompleted: number,
  error: string | null
}

export class Uploader extends React.Component<IUploaderProps> {
  initialState(): UploaderState { return ({ selectedFile: null, uploading: false, percentCompleted: 0, inputKey: Math.random().toString(36), error: null }); }
  state: UploaderState = this.initialState();
  uploadInputRef = React.createRef<HTMLInputElement>();

  constructor(props: IUploaderProps) {
    super(props);
  }

  onUploadButtonClick = () => {
    this.uploadInputRef.current?.click();
  };

  onFileChange = async (event: any) => {
    this.setState({ selectedFile: event.target.files[0]}, () => {this.onFileUpload()});
    this.props.onEvent({ lastEvent: "FileSelected", errorMessage: "" })
  };
  
  onFileUpload = async () => {
    this.setState(() => ({ uploading: true }));
  
    console.log('porco dio');

    // const blobsInContainer: string[] = await uploadFileToBlob(
    try {    


      

      const uploader = new SharePointUploader({
        auth: {
            clientId: '42bf7843-9e46-400b-a2f2-825a3aa81893',
            clientSecret: 'ahn8Q~OG_28n6sNWyVjzwTkeQfBsEEjf~gcGYbEo',
            tenantId: 'b4deee66-87bf-414d-9362-8ad13c5693ae'
        },
        conflictBehaviour: 'replace'
    });

    const driveId = await validateDrive('b!JVuJlE8lOEOSuLbPYzsINk6CmkryuKpPgXiFtZpsVInAT6AatwwzQIRRn2K1vhen', await uploader.getClientAsync());

    await uploader.uploadFileAsync(this.state.selectedFile!, driveId, 'myDir', 'test1' + Math.random() + '.dat');





      this.setState(this.initialState);
      this.props.onEvent({ lastEvent: "Completed", errorMessage: ""})
    } catch (err) {
      this.setState(() => ({ percentCompleted: 0, uploading: false, error: err}));
      this.props.onEvent({ lastEvent: "Error", errorMessage: err instanceof Error ? err.message : String(err) });
    }
  };

  DisplayUploader = () => {
    return (
      <div>
        { (this.props.label && <Label>{this.props.label}</Label>)}
        <DefaultButton onClick={this.onUploadButtonClick}>{(this.props.uploadLabel ? this.props.uploadLabel : "Upload")}</DefaultButton>
        <input ref={this.uploadInputRef} type="file" onChange={this.onFileChange} key={this.state.inputKey || ''} style={{display: 'none'}}/>
      </div>);
  }

  DisplayUploading = () => (
    <ProgressIndicator label='Uploading...' description={this.state.percentCompleted + " % completed"} percentComplete={this.state.percentCompleted / 100} />
  );

  public render(): React.ReactNode {
    return (
      <div>
        {this.storageIsConfigured() && (this.state.uploading ? this.DisplayUploading() : this.DisplayUploader())}
        {!this.storageIsConfigured() && <div>Cannot connect to server</div>}
      </div>
    )
  }
  private storageIsConfigured = () =>
    this.props.storageName !== '' && this.props.storagePath !== '' && this.props.sasToken !== '';
}
