/* eslint-disable react/jsx-key */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable dot-notation */
import * as React from 'react';
import styles from './HdrCandidate.module.scss';
import { IHdrCandidateProps } from './IHdrCandidateProps';
import IHdrCandidateState from './IHdrCandidateState';
import ScreenState from "./ScreenState";

export default class HdrCandidate extends React.Component<IHdrCandidateProps, IHdrCandidateState> {

  uploadFileChange(e:any):void{
    const allFiles:File[] = this.state?.myFiles;
    if(e.target.files){
      for(let i=0;i<e.target.files.length;i++){
        let hasFile=false;
        allFiles.map((myFile)=>{if(e.target.files[i].name===myFile.name){hasFile = true;}})
        if(!hasFile){
          allFiles.push(e.target.files[i]);
        }
      }
    }
    this.setState({myFiles:allFiles});
  }

  constructor(props: IHdrCandidateProps){
    super(props);
    this.state = {
      myFiles:[],
      screenState:ScreenState.submit
    };
  }

  popFile(e:any):void{
    const idToPop = parseInt(e.target.attributes['itemid'].value);
    const allFiles:File[] = this.state?.myFiles;
    const newFiles:File[] = [];
    for(let i=0;i<allFiles.length;i++)
    {
      if(i!==idToPop)
      {
        newFiles.push(allFiles[i]);
      }
    }
    this.setState({myFiles:newFiles});
  }

  async getFileBuffer(files: File[], i: number): Promise<any> {
    return new Promise((resolve, reject) => {
      // eslint-disable-next-line prefer-const
      let reader = new FileReader();
      reader.onloadend = function (e) {
        resolve(e.target?.result);
      }
      reader.onerror = function (e) {
        reject(e.target?.error);
      }
      reader.readAsArrayBuffer(files[i]);
    })
  }

  async addFileToFolder(files: File[], formDigest: any, arrayBuffer: any, i: number, id: number): Promise<any> {
    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve,reject) => {
      // Get the file name from the file input control on the page.
      const fileName = files[i].name;

      // Send the request and return the response.
      // This call returns the SharePoint file.
      await this.props.ctx.httpClient.post(this.props.ctx.pageContext.web.serverRelativeUrl + '/_api/Lists/getByTitle(\'Candidate Attachments\')/Items(' + id + ')/AttachmentFiles/add(FileName=\'' + fileName + '\')',undefined,{body:
        arrayBuffer,headers: {
            'X-RequestDigest': formDigest,
            //'content-length': '\''+arrayBuffer.byteLength+'\'',
            accept: 'application/json;odata=verbose'
          }
        }
      ).then((t:any) => {
        t.json().then((y:any)=>resolve(y));
      },y=>reject(y));
    })
  }

  async uploadAttachments(files: File[], formDigest: any, id: number, filesUploaded: number): Promise<any> {
    // eslint-disable-next-line no-async-promise-executor
    return new Promise(async (resolve, reject) => {
      try {
        const fileCount = files.length;
         await this.getFileBuffer(files, filesUploaded).then(async arrayBuffer => {
          await this.addFileToFolder(files, formDigest, arrayBuffer, filesUploaded, id).then(async value => {
            filesUploaded++;
            if (fileCount === filesUploaded) {
              resolve(value);
            }
            else {
              await this.uploadAttachments(files, formDigest, id, filesUploaded).then(t => resolve(t),y=>reject(y));
            }
          },(value)=>reject(value))
        },(value)=>reject(value));
      } catch (ex) {
        reject(ex);
      }
    });
  }

  async onSubmit():Promise<boolean>{
    this.setState({screenState:ScreenState.uploading});
    const url = new URL(window.location.href);
    const params = url.searchParams;
    const spListItemID = params.get('SID')!==null?parseInt(params.get('SID')):params.get('sid')!==null?parseInt(params.get('sid')):parseInt(params.get('Sid'));
    
    console.debug(spListItemID);
    console.debug(this.props.ctx.pageContext.web.serverRelativeUrl);

    await this.props.ctx.httpClient.post(this.props.ctx.pageContext.web.serverRelativeUrl+ '/_api/contextinfo',undefined,{body:'',headers:{accept:'application/json'}}).then(async (clientResponse)=>{
      await clientResponse.json().then(async t=>{
        await this.uploadAttachments(this.state.myFiles,t.FormDigestValue,spListItemID,0).then(()=>this.setState({screenState:ScreenState.success}),()=>{this.setState({screenState:ScreenState.error})})
      })
    },(clientReject)=>{console.debug(clientReject)});
    return false;
  }

  public render(): any {
    const {
      hasTeamsContext
    } = this.props;
    
    return this.state.screenState===ScreenState.submit?(
      <section className={`${styles.hdrCandidate} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className="logoDiv"><img src="https://digital-standards.acu.edu.au/img/acu-logo-positive.svg" className="logo"/></div>
        <div><h1>Candidate file upload</h1><span className="type-tag--bottom"/></div>
        <form id="candidateUploadForm">
          <div className="TableRow">
            <div className="tableText">Attachments</div>
            <div className="TableColumn">
            As per Section 10.5 of the <a href="https://policies.acu.edu.au/student_policies/higher_degree_research_regulations">Higher Degree Research Regulations and procedures</a>, you need to provide:<br/>
            - Two chapters (clean)<br/>
            - Two chapters (with iThenticate report)<br/>
            - Progress report<br/>
            - any other documents such as a copy of presentation slides (optional).<br/><br/>
            If you have any questions about milestone requirements, please speak with your Principal Supervisor or Associate Dean Research (ADR).<br /><br />
            Please note that you will require to remove existing files (using the &#39;X&#39;) and add again to upload a new version with the same name. Add file will not overwrite your files.<br/><br/>
              <label htmlFor='file' className="button primary">Add File(s)</label>
              <input id="file" className='fileUploadOpacity' type="file" onChange={this.uploadFileChange.bind(this)}
                multiple accept="image/*,.pdf,.doc,.docx,.pptx" />
              &nbsp;<br />
              <div>
              <table>
                <th/>
                <th>File Name</th>
                 {
                  this.state.myFiles.map(
                    (uploadedFile,index)=>{
                      return <tr>
                                <td itemID={index.toString()} onClick={this.popFile.bind(this)} className='popFile'>X</td>
                                <td>{uploadedFile.name}</td>
                              </tr>
                    }
                  )
                }
              </table>
            </div>
            </div>
          </div><br/>
          {this.state.myFiles.length>0?<span id="SubmitTag" className="button primary" onClick={this.onSubmit.bind(this)}>Submit</span>:<br/>}
        </form>
      </section>
    )
    :
    this.state.screenState===ScreenState.success
    ?
    <section className={`${styles.hdrCandidate} ${hasTeamsContext ? styles.teams : ''}`}>
    <div className="logoDiv"><img src="https://digital-standards.acu.edu.au/img/acu-logo-positive.svg" className="logo"/></div>
      <div>
        <h1>Candidate file upload</h1>
        <span className="type-tag--bottom"/>
      </div><br/>
      <form id="candidateUploadForm">
        <div className="TableRow">
          <div className="tableText">File Upload Success</div>
          <div className="TableColumn">Your files have been successfully uploaded. <br/>
          Your principal supervisor would now get an approval task to review and approve your attachments.<br/>
          Ppon their approval, the documents would be communicated to all the attendees.<br/>
          In case of their rejection, you would be contacted via email for the same to upload the attachments again.</div>
        </div>
      </form>
    </section>
    :
    this.state.screenState===ScreenState.uploading
    ?
    <section className={`${styles.hdrCandidate} ${hasTeamsContext ? styles.teams : ''}`}>
    <div className="logoDiv"><img src="https://digital-standards.acu.edu.au/img/acu-logo-positive.svg" className="logo"/></div>
      <div>
        <h1>Candidate file upload</h1>
        <span className="type-tag--bottom"/>
      </div><br/>
      <div>Files Upload in progress..</div>
    </section>
    :
    <section className={`${styles.hdrCandidate} ${hasTeamsContext ? styles.teams : ''}`}>
    <div className="logoDiv"><img src="https://digital-standards.acu.edu.au/img/acu-logo-positive.svg" className="logo"/></div>
      <div>
        <h1>Candidate file upload</h1>
        <span className="type-tag--bottom"/>
      </div><br/>
      <div>File Upload Error</div>
    </section>;
  }
}

