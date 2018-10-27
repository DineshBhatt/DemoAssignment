import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as JQuery from 'jquery';
import { SPComponentLoader } from '@microsoft/sp-loader'
require('bootstrap');
import styles from './DemoAssignmentWebPart.module.scss';
import * as strings from 'DemoAssignmentWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IDemoAssignmentWebPartProps {
  description: string;
}

export default class DemoAssignmentWebPart extends BaseClientSideWebPart<IDemoAssignmentWebPartProps> {
  
  public render(): void {
    
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css");
    this.domElement.innerHTML = `
      <div class="${ styles.demoAssignment }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
      <div class="Location"></div>`;
      this.getLocationInforamtion();
      JQuery(document).ready(function (){});
  }
  getLocationInforamtion(){
    let LocVar:string ='';
    if(Environment.type===EnvironmentType.Local){
      this.domElement.querySelector('.Location').innerHTML = "no Location found";
    }else{
      this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists/getByTitle('DineshVotingLocation')/items?$select=Title,Image,ID,Locations", SPHttpClient.configurations.v1
      ).then((Respons : SPHttpClientResponse)=>{
        Respons.json().then((listsObjects: any)=>{
          listsObjects.value.forEach(element => {
            LocVar += `<div class='col-md-3' ><img src="${element.Image}" alt="${element.Title}" style="width:100%; height:100%"><br /><h1>${element.Locations}</h1><br /><button class='btn btn-primary' type="button" id="${element.ID}" /><i class='fa fa-thumbs-up fa-2x'></div>`
          });
          this.domElement.querySelector('.Location').innerHTML = LocVar;
        });
      });
    }
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
