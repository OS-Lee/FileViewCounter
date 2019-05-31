import * as React from 'react';
import styles from './FileViewCounter.module.scss';
import { IFileViewCounterProps } from './IFileViewCounterProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IFileViewData{ 
  FileViewResult:string;
}

export default class FileViewCounter extends React.Component<IFileViewCounterProps, IFileViewData> {

  public constructor(props: IFileViewCounterProps,state: IFileViewData){ 
    super(props); 
    this.state = { 
      FileViewResult:""      
    }; 
  }

  public componentDidMount(){
    if(this.props.listName!=null && this.props.listName!=""){
      console.log('componentDidMount');
      var reactHandler = this;
      this.loadRootDrive().then(data=>{      
        reactHandler.setState({ 
          FileViewResult:data
        });           
        
        SPComponentLoader.loadScript('https://code.jquery.com/jquery-2.1.1.min.js', {
              globalExportsName: 'jQuery'
            }).then(($: any) => {
              var jQuery = $;
              SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery-treegrid/0.2.0/js/jquery.treegrid.min.js', {
                globalExportsName: 'jQuery'
              }).then(() => {
                jQuery(".tree").treegrid();   
              });
            });           
      });
    }    
  }

  private loadDriveView(Item:any,IsRoot:boolean): Promise<string> {
    return new Promise<string>((resolve: (data: string) => void, reject: (error: any) => void) => { 
    var dataString:string;
    if(Item.folder!=null){
      if(IsRoot){ 
        dataString="<tr class='treegrid-"+Item.id+"'>"+
                        "<td>"+Item.name+"</td>"+
                        "<td></td>"+
                        "<td></td>"+
                      "</tr>";   
        }else{
          dataString="<tr class='treegrid-"+Item.id+" treegrid-parent-"+Item.parentReference.id+"'>"+
                        "<td>"+Item.name+"</td>"+
                        "<td></td>"+
                        "<td></td>"+
                      "</tr>"; 
        }         
      this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/v2.1/drives/`+this.props.listName+`/items/`+Item.id+`/children`,
      SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => {         
          Promise.all(responseJSON.value.map(item => this.loadDriveView(item,false))).then((results:any[]) =>
          {
            results.forEach((element) => { 
              dataString+=element;
            });
            resolve(dataString);
          });           
        });  
      });
    }else{
      this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/v2.1/drives/`+this.props.listName+`/items/`+Item.id+`?select=id,name,parentReference,dataLossPrevention&$expand=analytics($expand=allTime($expand=activities($filter=access%20ne%20null)))`,
      SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => { 
          if(IsRoot){ 
            if(responseJSON.analytics.allTime.access!=null&&responseJSON.analytics.allTime.access.actionCount!=null){
              dataString="<tr class='treegrid-"+Item.id+"'>"+
                              "<td>"+responseJSON.name+"</td>"+
                              "<td>"+responseJSON.analytics.allTime.access.actionCount+"</td>"+
                              "<td>"+responseJSON.analytics.allTime.access.actorCount+"</td>"+
                            "</tr>";  
            }else{
              dataString="<tr class='treegrid-"+Item.id+"'>"+
                              "<td>"+responseJSON.name+"</td>"+
                              "<td>retrieve actionCount failed</td>"+
                              "<td>retrieve actorCount failed</td>"+
                            "</tr>";
            } 
          }else{             
            if(responseJSON.analytics.allTime.access!=null&&responseJSON.analytics.allTime.access.actionCount!=null){
              dataString="<tr class='treegrid-"+Item.id+" treegrid-parent-"+Item.parentReference.id+"'>"+
                          "<td>"+responseJSON.name+"</td>"+
                          "<td>"+responseJSON.analytics.allTime.access.actionCount+"</td>"+
                          "<td>"+responseJSON.analytics.allTime.access.actorCount+"</td>"+
                        "</tr>";    
            }else{
              //console.log(responseJSON);              
              dataString="<tr class='treegrid-"+Item.id+" treegrid-parent-"+Item.parentReference.id+"'>"+
                          "<td>"+responseJSON.name+"</td>"+
                          "<td>retrieve actionCount failed</td>"+
                          "<td>retrieve actorCount failed</td>"+
                        "</tr>";
            }       
          }
          resolve(dataString);                 
        });  
      });
      
    }    
    });            
  }

  private loadRootDrive(): Promise<string> {
    return new Promise<string>((resolve: (data: string) => void, reject: (error: any) => void) => { 
    var htmlRender="";     
    this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/v2.1/drives/`+this.props.listName+`/root/children`,  
      SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => {                     
          Promise.all(responseJSON.value.map(item => this.loadDriveView(item,true))).then((results:any[]) =>
          {
            results.forEach((element) => { 
              htmlRender+=element;
            });
            resolve(htmlRender);
          });                     
        });  
      });
    });
  }

  public render(): React.ReactElement<IFileViewCounterProps> {      
    if(this.props.listName==null || this.props.listName==""){
      return(<div className={ styles.fileViewCounter }>
      <div className={ styles.container }>
        <div className={ styles.row }>
          <div className={ styles.column }>              
            configure List Property first.         
          </div>
        </div>
      </div>
    </div>
  ); 
    }else{  
    return (      
      <div className={ styles.fileViewCounter }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>              
              <table className="tree">
                <thead>
                    <tr>
                        <th className={styles.THWidth}>Folder/File</th>
                        <th >Views</th>
                        <th >Viewers</th>
                    </tr>
                </thead>                
                { this.state.FileViewResult
                    ? <tbody dangerouslySetInnerHTML={{ __html: this.state.FileViewResult}}></tbody>
                    : <tr><td align={"center"} colSpan={3}>Data Loading...</td></tr>
                 }                
            </table>          
            </div>
          </div>
        </div>
      </div>
    ); 
  }
  }
}
