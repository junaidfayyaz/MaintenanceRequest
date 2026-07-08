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
import * as React from "react";
import { useEffect, useState } from "react";
import AlertModal from "../../../components/alertModal/AlertModal";
import MaintenanceRequestUIForm from "../../../components/MaintenanceRequestUIForm";
var isAr = window.location.pathname.includes("/ar/") ||
    window.location.search.includes("lang=ar");
var MaintenanceRequest = function (props) {
    var _a = useState(null), userProfileAD = _a[0], setUserProfileAD = _a[1];
    var _b = useState(true), isLoadingUser = _b[0], setIsLoadingUser = _b[1];
    var _c = useState(false), showModal = _c[0], setShowModal = _c[1];
    var _d = useState(""), modalHeading = _d[0], setModalHeading = _d[1];
    var _e = useState(""), modalMessage = _e[0], setModalMessage = _e[1];
    var _f = useState(""), alertsection = _f[0], setAlertsection = _f[1];
    var _g = useState(""), iconLoad = _g[0], setIconLoad = _g[1];
    var handleShowModal = function () { return setShowModal(true); };
    var handleCloseModal1 = function (section) {
        if (section === "Accepted") {
            setShowModal(false);
            window.location.reload();
        }
        else {
            setShowModal(false);
        }
    };
    useEffect(function () {
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var client, userAD, userProfile, error_1;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, props.context.msGraphClientFactory.getClient("3")];
                    case 1:
                        client = _b.sent();
                        return [4 /*yield*/, client
                                .api("/me")
                                .select("displayName,jobTitle,department,employeeId,mail,onPremisesExtensionAttributes")
                                .get()];
                    case 2:
                        userAD = _b.sent();
                        userProfile = {
                            displayName: userAD.displayName || "",
                            jobTitle: userAD.jobTitle || "",
                            department: userAD.department || "",
                            mail: userAD.mail || "",
                            employeeId: ((_a = userAD === null || userAD === void 0 ? void 0 : userAD.onPremisesExtensionAttributes) === null || _a === void 0 ? void 0 : _a.extensionAttribute15) || "",
                        };
                        setUserProfileAD(userProfile);
                        setIsLoadingUser(false);
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _b.sent();
                        console.error("Error fetching user info:", error_1);
                        setIsLoadingUser(false);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); })();
    }, [props]);
    var showErrorModal = function () {
        setModalHeading("Warning");
        setModalMessage("Please fill Required fields");
        setAlertsection("rejected");
        setIconLoad("WarningSolid");
        handleShowModal();
    };
    var saveRequest = function (formData) { return __awaiter(void 0, void 0, void 0, function () {
        var params, payload, response, errorText, rawResponse, jsonStart, jsonString, parsedData, dataObj, rpcData, strRequestNum_1, requestRecId_1, flag, error_2;
        var _a;
        var _b, _c, _d;
        return __generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    _e.trys.push([0, 5, , 6]);
                    params = (_a = {},
                        _a[props.Category_col] = formData.Category,
                        _a[props.SubCategory_col] = formData.SubCategory,
                        _a[props.Department_col] = formData.Department,
                        _a[props.Location_col] = formData.Location,
                        _a[props.Description_col] = formData.Description,
                        _a[props.ProfileLink_col] = formData.requestedBy_key || formData.requestedFor_key,
                        _a[props.SubmittedBy_col] = formData.requestedBy,
                        _a);
                    if (formData.Category_key) {
                        params["".concat(props.Category_col, "-recId")] = formData.Category_key;
                    }
                    if (formData.SubCategory_key) {
                        params["".concat(props.SubCategory_col, "-recId")] = formData.SubCategory_key;
                    }
                    payload = {
                        attachmentsToDelete: [],
                        attachmentsToUpload: [],
                        parameters: params,
                        delayedFulfill: false,
                        formName: "ServiceReq.ResponsiveAnalyst.DefaultLayout",
                        saveReqState: false,
                        serviceReqData: {
                            Subject: formData.Category ? "Maintenance Request - ".concat(formData.Category) : "Maintenance Request",
                            Symptom: formData.Description ? formData.Description : "Maintenance Request",
                            Category: "",
                            CreatedBy: formData.requestedBy,
                        },
                        subscriptionId: "EE937A037A90490D8A38B2437B567673",
                    };
                    console.log("=== VERIFY DATA ===");
                    console.log("Raw Form Data:", formData);
                    console.log("API Payload:", payload);
                    console.log("===================");
                    return [4 /*yield*/, fetch(props.Apilink, {
                            method: "POST",
                            headers: {
                                "Content-Type": "application/json",
                                accept: "application/json;odata=verbose",
                                "Ocp-Apim-Subscription-Key": props.OcpApimKey,
                                Authorization: props.Authorization,
                                Email: formData.requestedFor,
                            },
                            body: JSON.stringify(payload),
                        })];
                case 1:
                    response = _e.sent();
                    if (!!response.ok) return [3 /*break*/, 3];
                    return [4 /*yield*/, response.text()];
                case 2:
                    errorText = _e.sent();
                    throw new Error("Request failed: ".concat(response.status, " - ").concat(errorText));
                case 3: return [4 /*yield*/, response.text()];
                case 4:
                    rawResponse = _e.sent();
                    jsonStart = rawResponse.indexOf("{");
                    if (jsonStart === -1) {
                        throw new Error("JSON not found in response");
                    }
                    jsonString = rawResponse.slice(jsonStart);
                    parsedData = void 0;
                    try {
                        parsedData = JSON.parse(jsonString);
                    }
                    catch (e) {
                        throw new Error("Failed to parse JSON: " + e.message);
                    }
                    dataObj = Array.isArray(parsedData) ? parsedData[0] : parsedData;
                    rpcData = ((_c = (_b = dataObj === null || dataObj === void 0 ? void 0 : dataObj.result) === null || _b === void 0 ? void 0 : _b.ServiceRequests) === null || _c === void 0 ? void 0 : _c[0]) || ((_d = dataObj === null || dataObj === void 0 ? void 0 : dataObj.ServiceRequests) === null || _d === void 0 ? void 0 : _d[0]);
                    strRequestNum_1 = (rpcData === null || rpcData === void 0 ? void 0 : rpcData.strRequestNum) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.IncidentNumber) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.ServiceReqNumber) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.requestNumber) || (parsedData === null || parsedData === void 0 ? void 0 : parsedData.strRequestNum) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.ServiceReqNum);
                    requestRecId_1 = (rpcData === null || rpcData === void 0 ? void 0 : rpcData.strRequestRecId) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.RecId) || (dataObj === null || dataObj === void 0 ? void 0 : dataObj.recId) || (parsedData === null || parsedData === void 0 ? void 0 : parsedData.RecId);
                    console.log("=== API RESPONSE RESULT ===");
                    console.log("Full Response Data:", parsedData);
                    console.log("Extracted ID (strRequestNum):", strRequestNum_1);
                    console.log("Extracted RecId:", requestRecId_1);
                    console.log("===========================");
                    flag = true;
                    if (formData.files && formData.files.length > 0) {
                        flag = false;
                        formData.files.forEach(function (element) { return __awaiter(void 0, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, saveRequestAttachment(requestRecId_1, strRequestNum_1, element, formData.requestedFor)];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); });
                    }
                    if (flag) {
                        setModalHeading(isAr ? "تمت العمليه بنجاح" : "Success");
                        setModalMessage(isAr
                            ? "\u062A\u0645 \u062A\u0642\u062F\u064A\u0645 \u0637\u0644\u0628\u0643 \u0628\u0646\u062C\u0627\u062D. \u0631\u0642\u0645 \u0627\u0644\u062A\u0639\u0631\u064A\u0641:".concat(strRequestNum_1)
                            : "Your Request has been submitted successfully. ID:".concat(strRequestNum_1));
                        setAlertsection("Accepted");
                        setIconLoad("SkypeCircleCheck");
                        handleShowModal();
                    }
                    return [3 /*break*/, 6];
                case 5:
                    error_2 = _e.sent();
                    console.error("Error submitting Request:", error_2);
                    setModalHeading("Error");
                    setModalMessage(error_2.message);
                    setAlertsection("rejected");
                    setIconLoad("ErrorBadge");
                    handleShowModal();
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    }); };
    var saveRequestAttachment = function (recid, requestnum, fileObj, userEmail) { return __awaiter(void 0, void 0, void 0, function () {
        var ApiformData, apiUrl, response, errorText, error_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 5, , 6]);
                    ApiformData = new FormData();
                    ApiformData.append("ObjectID", recid);
                    ApiformData.append("ObjectType", "Incident#"); // Based on API, might need to change if using ServiceReq
                    ApiformData.append("File", fileObj.content, fileObj.name);
                    apiUrl = props.attachmentApilink;
                    return [4 /*yield*/, fetch(apiUrl, {
                            method: "POST",
                            headers: {
                                accept: "application/json",
                                "Ocp-Apim-Subscription-Key": props.OcpApimKey,
                                Authorization: props.Authorization,
                                Email: userEmail,
                            },
                            body: ApiformData,
                        })];
                case 1:
                    response = _a.sent();
                    if (!response.ok) return [3 /*break*/, 2];
                    setModalHeading(isAr ? "تمت العمليه بنجاح" : "Success");
                    setModalMessage(isAr
                        ? "\u062A\u0645 \u062A\u0642\u062F\u064A\u0645 \u0637\u0644\u0628\u0643 \u0628\u0646\u062C\u0627\u062D. \u0631\u0642\u0645 \u0627\u0644\u062A\u0639\u0631\u064A\u0641:".concat(requestnum)
                        : "Your Request has been submitted successfully. ID:".concat(requestnum));
                    setAlertsection("Accepted");
                    setIconLoad("SkypeCircleCheck");
                    handleShowModal();
                    return [3 /*break*/, 4];
                case 2: return [4 /*yield*/, response.text()];
                case 3:
                    errorText = _a.sent();
                    throw new Error("Request failed: ".concat(response.status, " - ").concat(errorText));
                case 4: return [3 /*break*/, 6];
                case 5:
                    error_3 = _a.sent();
                    console.error("Error submitting Attachment:", error_3);
                    setModalHeading("Error");
                    setModalMessage(error_3.message);
                    setAlertsection("rejected");
                    setIconLoad("ErrorBadge");
                    handleShowModal();
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    }); };
    if (isLoadingUser) {
        return React.createElement("div", null, "Loading user information...");
    }
    return (React.createElement(React.Fragment, null,
        React.createElement(MaintenanceRequestUIForm, { Authorization: props.Authorization, OcpApimKey: props.OcpApimKey, UserRecIdApilink: props.UserRecIdApilink, context: props.context, userprofileAD: userProfileAD, EmpId: (userProfileAD === null || userProfileAD === void 0 ? void 0 : userProfileAD.employeeId) || "", CurrentUserEmail: props.CurrentUserEmail || "", onErrorRequiredFields: function () { return showErrorModal(); }, onSave: function (formData) { return __awaiter(void 0, void 0, void 0, function () {
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0: return [4 /*yield*/, saveRequest(formData)];
                        case 1:
                            _a.sent();
                            return [2 /*return*/];
                    }
                });
            }); } }),
        React.createElement(AlertModal, { showModal: showModal, handleShowModal: handleShowModal, handleCloseModal: handleCloseModal1, heading: modalHeading, message: modalMessage, style: "", section: alertsection, icon: iconLoad, isAr: isAr })));
};
export default MaintenanceRequest;
//# sourceMappingURL=MaintenanceRequest.js.map