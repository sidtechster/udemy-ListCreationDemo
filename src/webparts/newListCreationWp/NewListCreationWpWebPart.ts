import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewListCreationWpWebPart.module.scss';
import * as strings from 'NewListCreationWpWebPartStrings';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface INewListCreationWpWebPartProps {
  description: string;
}

export default class NewListCreationWpWebPart extends BaseClientSideWebPart<INewListCreationWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.newListCreationWp }">
        
        <h3>Creating a new list dynamically</h3><br/><br/><br/>

        <p>Please fill out the below details</p><br/><br/>

        New list name: <br/><input type='text' id='txtNewListName' /><br/><br/>

        New list description: <br/><input type='text' id='txtNewListDesc' /><br/><br/>

        <input type='button' id='btnCreateNewList' value='Create new list'/></br>

      </div>`;

      this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnCreateNewList').addEventListener('click', () => { this.createNewList(); });
  }

  private createNewList(): void {

    var newListName = document.getElementById('txtNewListName')["value"];
    var newListDesc = document.getElementById('txtNewListDesc')["value"];

    const listUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + newListName + "')";

    this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 200) {
          alert('A list already exists with that name');
          return;
        }
        if(response.status === 404) {
          const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
          const listDefinition: any = {
            "Title": newListName,
            "Description": newListDesc,
            "AllowContentTypes": true,
            "BaseTemplate": 100,
            "ContentTypesEnabled": true
          };
          const spHttpClientOptions: ISPHttpClientOptions = {
            "body": JSON.stringify(listDefinition)
          };
          this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((response: SPHttpClientResponse) => {
              if(response.status === 201) {
                alert('List created successfully');
              }
              else {
                alert('Error: ' + response.status + ' - ' + response.statusText);
              }
            });
        }
      })

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
