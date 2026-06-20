var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ServiceWebPartStrings';
import MaintenanceRequest from './components/MaintenanceRequest';
import { sp } from "@pnp/sp";
var MaintenanceRequestWebPart = /** @class */ (function (_super) {
    __extends(MaintenanceRequestWebPart, _super);
    function MaintenanceRequestWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    MaintenanceRequestWebPart.prototype.render = function () {
        var element = React.createElement(MaintenanceRequest, {
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
        });
        ReactDom.render(element, this.domElement);
    };
    MaintenanceRequestWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                sp.setup({
                    sp: {
                        baseUrl: this.context.pageContext.web.absoluteUrl
                    }
                });
                // Default configuration values
                if (!this.properties.Apilink) {
                    this.properties.Apilink = "https://mcitstgapim.mcit.gov.qa/ITSM/newIncident";
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
                if (!this.properties.Category_col)
                    this.properties.Category_col = "par-2D343B2215B945A88B82D5299CA4536D";
                if (!this.properties.SubCategory_col)
                    this.properties.SubCategory_col = "par-C06F7C725B924E09ABF2D8EBBB103CA8";
                if (!this.properties.Department_col)
                    this.properties.Department_col = "par-683F9D91D9D54A03B795551575E97AE1";
                if (!this.properties.Location_col)
                    this.properties.Location_col = "par-89C9D84DF7054849AEFF75D680008086";
                if (!this.properties.Description_col)
                    this.properties.Description_col = "par-08A9F3D288354138B6A8018622277C43";
                if (!this.properties.ProfileLink_col)
                    this.properties.ProfileLink_col = "ProfileLink";
                if (!this.properties.SubmittedBy_col)
                    this.properties.SubmittedBy_col = "SubmittedBy";
                return [2 /*return*/, _super.prototype.onInit.call(this)];
            });
        });
    };
    MaintenanceRequestWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return MaintenanceRequestWebPart;
}(BaseClientSideWebPart));
export default MaintenanceRequestWebPart;
//# sourceMappingURL=MaintenanceRequestWebPart.js.map