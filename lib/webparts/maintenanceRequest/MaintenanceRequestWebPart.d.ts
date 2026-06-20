import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
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
    render(): void;
    protected onInit(): Promise<void>;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=MaintenanceRequestWebPart.d.ts.map