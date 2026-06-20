import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../css/style.css";
import { IUserProfile, IMaintenanceRequestFormData } from "../webparts/maintenanceRequest/components/IMaintenanceRequestProps";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface IRequestUIFormProps {
    context: WebPartContext;
    userprofileAD: IUserProfile;
    EmpId: string;
    onErrorRequiredFields: () => void;
    onSave: (formData: IMaintenanceRequestFormData) => Promise<void>;
    OcpApimKey: string;
    UserRecIdApilink: string;
    Authorization: string;
    CurrentUserEmail?: string;
}
declare const MaintenanceRequestUIForm: React.FC<IRequestUIFormProps>;
export default MaintenanceRequestUIForm;
//# sourceMappingURL=MaintenanceRequestUIForm.d.ts.map