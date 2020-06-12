import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { IListItem } from './IListItem';
import styles from './SpfxcrudWebPart.module.scss';
import * as strings from 'SpfxcrudWebPartStrings';

export interface ISpfxcrudWebPartProps {
  listName: string;
}

export default class SpfxcrudWebPart extends BaseClientSideWebPart <ISpfxcrudWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spfxcrud }">
    <div class="${ styles.container }">
      <div class="${ styles.row }">
        <div class="${ styles.column }">
        <span class="${ styles.title }">CRUD operations</span>  
        <p class="${ styles.subTitle }">No Framework</p>  
        <p class="${ styles.description }">Name: ${escape(this.properties.listName)}</p> 

        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">   
            <span class="${styles.label}">Title</span>  
            <input type="text" id="Title"/> 
           <br/><br/>
            <span class="${styles.label}">Test</span>  
            <input type="text" id="test"/> 

        </div>  
      </div> 

        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
            <button class="${styles.button} create-Button">  
              <span class="${styles.label}">Create item</span>  
            </button>  
            <button class="${styles.button} read-Button">  
              <span class="${styles.label}">Read item</span>  
            </button>  
          </div>  
        </div>  

        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
            <button class="${styles.button} update-Button">  
              <span class="${styles.label}">Update item</span>  
            </button>  
            <button class="${styles.button} delete-Button">  
              <span class="${styles.label}">Delete item</span>  
            </button>  
          </div>  
        </div>  

        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
            <div class="status"></div>  
            <ul class="items"><ul>  
            <div id="listitems"></div>
          </div>  
        </div>  

      </div>  
    </div>  
  </div>  
</div>`;  

this.setButtonsEventHandlers();  
}

private setButtonsEventHandlers(): void {  
  const webPart: SpfxcrudWebPart = this;  
  this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });  
  this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem().then((response) => {this._renderList(response.value); }); });  
  this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });  
  this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });  
}  

private createItem(): void {  
  const body: string = JSON.stringify({  
    'Title': `${document.getElementById("Title")["value"]}`,
    'test':` ${document.getElementById("test")["value"]}`  
  });  
  
  this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items`,  
  SPHttpClient.configurations.v1,  
  {  
    headers: {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': ''  
    },  
    body: body  
  })  
  .then((response: SPHttpClientResponse): Promise<IListItem> => {  
    return response.json();  
  })  
  .then((item: IListItem): void => {  
    this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);  
  }, (error: any): void => {  
    this.updateStatus('Error while creating the item: ' + error);  
  });  
}  
private updateStatus(status: string, items: IListItem[] = []): void {  
  this.domElement.querySelector('.status').innerHTML = status;  
  this.updateItemsHtml(items);  
}  
  
private updateItemsHtml(items: IListItem[]): void {  
  this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");  
}  
private readItem(): any {  
  this.getLatestItemId()  
    .then((itemId: number): Promise<SPHttpClientResponse> => {  
      if (itemId === -1) {  
        throw new Error('No items found in the list');  
      }  
        
      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$select=Title,Id,test`,  
        SPHttpClient.configurations.v1,  
        { 
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((response: SPHttpClientResponse): Promise<IListItem> => {  
      
      return response.json();  
    })
    this.updateStatus("data read complete");

  }
  private _renderList(items: IListItem[]): any {
    let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
    html += '<th>Title</th> <th>Product Code</th><th>Product Description</th>';
    items.forEach((item: IListItem) => {
      html += `
      <tr>            
          <td>${item.Id}</td>
          <td>${item.Title}</td>
          <td>${item.test}</td>
          
          </tr>
          `;
    });
    html += '</table>';
    this.domElement.querySelector('#listitems').innerHTML = html;
  }
private updateItem(): void {  
}  

private deleteItem(): void {  
}  

private getLatestItemId(): Promise<number> {  
  return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$orderby=Id desc&$top=1&$select=id`,  
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'odata-version': ''  
        }  
      })  
      .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
        return response.json();  
      }, (error: any): void => {  
        reject(error);  
      })  
      .then((response: { value: { Id: number }[] }): void => {  
        if (response.value.length === 0) {  
          resolve(-1);  
        }  
        else {  
          resolve(response.value[0].Id);  
        }  
      });  
  });  
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
              PropertyPaneTextField('listName', {
                label: strings.ListNameFieldLabel
              })
            ]
          }
        ]
      }
    ]
  };
}
}
