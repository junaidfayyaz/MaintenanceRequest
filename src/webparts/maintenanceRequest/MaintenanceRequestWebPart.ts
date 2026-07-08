import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ServiceWebPartStrings';
import MaintenanceRequest from './components/MaintenanceRequest';
import { IMaintenanceRequestProps } from './components/IMaintenanceRequestProps';
import { sp } from "@pnp/sp";

export interface IMaintenanceRequestWebPartProps {
  Apilink: string;
  attachmentApilink: string;
  UserRecIdApilink: string;
  OcpApimKey: string;
  Authorization: string;

  Category_col: string;
  SubCategory_col: string;
  Department_col: string;
  Location_col: string;
  Description_col: string;
  ProfileLink_col: string;
  SubmittedBy_col: string;
}

export default class MaintenanceRequestWebPart extends BaseClientSideWebPart<IMaintenanceRequestWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IMaintenanceRequestProps> = React.createElement(
      MaintenanceRequest,
      {
        context: this.context,
        Apilink: this.properties.Apilink,
        attachmentApilink: this.properties.attachmentApilink,
        UserRecIdApilink: this.properties.UserRecIdApilink,
        OcpApimKey: this.properties.OcpApimKey,
        Authorization: this.properties.Authorization,

        Category_col: this.properties.Category_col,
        SubCategory_col: this.properties.SubCategory_col,
        Department_col: this.properties.Department_col,
        Location_col: this.properties.Location_col,
        Description_col: this.properties.Description_col,
        ProfileLink_col: this.properties.ProfileLink_col,
        SubmittedBy_col: this.properties.SubmittedBy_col,

        CurrentUserEmail: this.context.pageContext.user.email
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    sp.setup({
      sp: {
        baseUrl: this.context.pageContext.web.absoluteUrl
      }
    });

    // Default configuration values
    if (!this.properties.Apilink) {
      this.properties.Apilink = "https://mcitstgapim.mcit.gov.qa/ivantitest/new";
    }
    if (!this.properties.attachmentApilink) {
      this.properties.attachmentApilink = "https://mcitstgapim.mcit.gov.qa/ivantitestattachment/newattachment";
    }
    if (!this.properties.UserRecIdApilink) {
      this.properties.UserRecIdApilink = "https://mcitstgapim.mcit.gov.qa/ITSM/employeeByEmail";
    }
    if (!this.properties.Authorization) {
      this.properties.Authorization = "api-key 21149AFFFF0046C0BCFEDAC94D46D452";
    }
    if (!this.properties.OcpApimKey) {
      this.properties.OcpApimKey = "ba47658772b3473cbd9eb045e856e9fc";
    }

    if (!this.properties.Category_col) this.properties.Category_col = "par-2D343B2215B945A88B82D5299CA4536D";
    if (!this.properties.SubCategory_col) this.properties.SubCategory_col = "par-0AF19ED8771E48D6B3A875AD546F131A";
    if (!this.properties.Department_col) this.properties.Department_col = "par-683F9D91D9D54A03B795551575E97AE1";
    if (!this.properties.Location_col) this.properties.Location_col = "par-89C9D84DF7054849AEFF75D680008086";
    if (!this.properties.Description_col) this.properties.Description_col = "par-08A9F3D288354138B6A8018622277C43";
    if (!this.properties.ProfileLink_col) this.properties.ProfileLink_col = "ProfileLink";
    if (!this.properties.SubmittedBy_col) this.properties.SubmittedBy_col = "SubmittedBy";

    return super.onInit();
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
              groupName: "API Configuration",
              groupFields: [
                PropertyPaneTextField('Apilink', {
                  label: "API Link",
                  value: this.properties.Apilink,
                }),
                PropertyPaneTextField('attachmentApilink', {
                  label: "Attachment API Link",
                  value: this.properties.attachmentApilink,
                }),
                PropertyPaneTextField('UserRecIdApilink', {
                  label: "User RecId API Link",
                  value: this.properties.UserRecIdApilink,
                }),
                PropertyPaneTextField('Authorization', {
                  label: "Authorization Id",
                  value: this.properties.Authorization,
                }),
                PropertyPaneTextField('OcpApimKey', {
                  label: "Ocp-Apim-Subscription-Key",
                  value: this.properties.OcpApimKey,
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Form Parameters"
          },
          groups: [
            {
              groupName: "Payload Mapping",
              groupFields: [
                PropertyPaneTextField('Category_col', {
                  label: "Category Parameter",
                  value: this.properties.Category_col,
                }),
                PropertyPaneTextField('SubCategory_col', {
                  label: "SubCategory Parameter",
                  value: this.properties.SubCategory_col,
                }),
                PropertyPaneTextField('Department_col', {
                  label: "Department Parameter",
                  value: this.properties.Department_col,
                }),
                PropertyPaneTextField('Location_col', {
                  label: "Location Parameter",
                  value: this.properties.Location_col,
                }),
                PropertyPaneTextField('Description_col', {
                  label: "Description Parameter",
                  value: this.properties.Description_col,
                }),
                PropertyPaneTextField('ProfileLink_col', {
                  label: "ProfileLink Parameter",
                  value: this.properties.ProfileLink_col,
                }),
                PropertyPaneTextField('SubmittedBy_col', {
                  label: "SubmittedBy Parameter",
                  value: this.properties.SubmittedBy_col,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}