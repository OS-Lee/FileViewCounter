import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';

import * as strings from 'FileViewCounterWebPartStrings';
import FileViewCounter from './components/FileViewCounter';
import { IFileViewCounterProps } from './components/IFileViewCounterProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IFileViewCounterWebPartProps {
  listName: string;
  description: string;
}

export default class FileViewCounterWebPart extends BaseClientSideWebPart<IFileViewCounterWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {    
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {      
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/v2.1/drives?select=id,name,drivetype`,  
      SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => {  
          var libraries=[];
          responseJSON.value.forEach(element => {
            libraries.push({key:element.id,text:element.name});
          });
          resolve(libraries);  
        });  
      });
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropdownDisabled = !this.lists;
    if (this.lists) {
      this.render();  
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
      });
  } 

  public constructor() {
    super();    
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/jquery-treegrid/0.2.0/css/jquery.treegrid.min.css");
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listName' && newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // re-render the web part as clearing the loading indicator removes the web part body
      this.render();      
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, oldValue);
    }
  }

  public render(): void {
    const element: React.ReactElement<IFileViewCounterProps > = React.createElement(
      FileViewCounter,
      {
        listName: this.properties.listName,
        description: this.properties.description,
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
    console.log("render function debug info"); 
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
