var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { TextField, PrimaryButton, DefaultButton, Dropdown } from "@fluentui/react";
import "../css/style.css";
import { PeoplePicker, PrincipalType, } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Col } from "react-bootstrap";
import AlertModal from "./alertModal/AlertModal";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
var isAr = window.location.pathname.includes("/ar/") ||
    window.location.search.includes("lang=ar");
var MaintenanceRequestUIForm = function (props) {
    var _a, _b, _c, _d, _e;
    var _f = useState({
        requestedBy: (_a = props.userprofileAD) === null || _a === void 0 ? void 0 : _a.displayName,
        requestedFor: "",
        requestedFor_Title: "",
        requestedFor_key: "",
        Category: "",
        SubCategory: "",
        Department: "",
        Location: "",
        Description: "",
        files: [],
    }), formData = _f[0], setFormData = _f[1];
    var _g = useState([]), uploadedFiles = _g[0], setUploadedFiles = _g[1];
    var _h = useState(""), showErrorUpload = _h[0], setShowErrorUpload = _h[1];
    var _j = useState(false), showSubmitbtn = _j[0], setShowSubmitbtn = _j[1];
    var _k = useState({}), errors = _k[0], setErrors = _k[1];
    var inputRef = useRef(null);
    var fileInputRef = React.useRef(null);
    var _l = useState(false), showEmptyFileModal = _l[0], setShowEmptyFileModal = _l[1];
    var _m = useState(0), setForceUpdater = _m[1];
    var _o = useState([]), categoryOptions = _o[0], setCategoryOptions = _o[1];
    var _p = useState([]), subCategoryOptions = _p[0], setSubCategoryOptions = _p[1];
    var _q = useState([]), listData = _q[0], setListData = _q[1];
    useEffect(function () {
        _getUserRecId(props.CurrentUserEmail, "requestedBy_key");
        fetchCategories();
    }, []);
    var fetchCategories = function () { return __awaiter(void 0, void 0, void 0, function () {
        var tahyahWeb, items, uniqueCategories, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    tahyahWeb = Web("https://mcitgovqa.sharepoint.com/sites/Tahyah");
                    return [4 /*yield*/, tahyahWeb.lists
                            .getByTitle("MaintenanceRequestCategoryList")
                            .items.getAll()];
                case 1:
                    items = _a.sent();
                    setListData(items);
                    uniqueCategories = Array.from(new Set(items.map(function (item) { return item.Category; }).filter(Boolean)));
                    setCategoryOptions(uniqueCategories.map(function (cat) { return ({ key: cat, text: cat }); }));
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    console.error("Error fetching categories: ", error_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleCategoryChange = function (event, option) {
        if (option) {
            var selectedCategory_1 = option.key;
            updateFormDropData(option, "Category");
            var categoryItem = listData.find(function (item) { return item.Category === selectedCategory_1; });
            var categoryKey_1 = "";
            if (categoryItem) {
                // Look for any property name that might contain the Category recId
                var possibleKeys = Object.keys(categoryItem).filter(function (k) { return k.toLowerCase().includes("recid") && !k.toLowerCase().includes("sub"); });
                if (possibleKeys.length > 0) {
                    categoryKey_1 = categoryItem[possibleKeys[0]];
                }
            }
            // Fallback to user-provided mapping if not found in SharePoint list
            if (!categoryKey_1 && selectedCategory_1 === "Access Control") {
                categoryKey_1 = "87967EC7EA98409C8C2A636924C6C797";
            }
            setFormData(function (prev) { return (__assign(__assign({}, prev), { Category_key: categoryKey_1, SubCategory: "", SubCategory_key: "" })); });
            // Filter subcategories based on selected category and ensure uniqueness
            var filteredSubCats = Array.from(new Set(listData
                .filter(function (item) { return item.Category === selectedCategory_1; })
                .map(function (item) { return item.SubCategory; })
                .filter(Boolean)));
            setSubCategoryOptions(filteredSubCats.map(function (subcat) { return ({ key: subcat, text: subcat }); }));
        }
    };
    var handleSubCategoryChange = function (event, option) {
        if (option) {
            var selectedSubCategory_1 = option.key;
            updateFormDropData(option, "SubCategory");
            var subCategoryItem_1 = listData.find(function (item) { return item.Category === formData.Category && item.SubCategory === selectedSubCategory_1; });
            var subCategoryKey_1 = "";
            if (subCategoryItem_1) {
                // Look for any property name that might contain the SubCategory recId
                var possibleKeys = Object.keys(subCategoryItem_1).filter(function (k) { return k.toLowerCase().includes("sub") && k.toLowerCase().includes("recid"); });
                if (possibleKeys.length > 0) {
                    subCategoryKey_1 = subCategoryItem_1[possibleKeys[0]];
                }
                else {
                    // fallback if it just named recId on the item level
                    var fallbackKeys = Object.keys(subCategoryItem_1).filter(function (k) { return k.toLowerCase().includes("recid"); });
                    if (fallbackKeys.length > 0) {
                        // pick one that is not the category key
                        var subKey = fallbackKeys.find(function (k) { return subCategoryItem_1[k] !== formData.Category_key; });
                        if (subKey)
                            subCategoryKey_1 = subCategoryItem_1[subKey];
                    }
                }
            }
            // Fallback to user-provided mapping if not found in SharePoint list
            if (!subCategoryKey_1 && selectedSubCategory_1 === "automated door noisy") {
                subCategoryKey_1 = "6BFF811DE98A4C17B21D38E084DEB1BB";
            }
            setFormData(function (prev) { return (__assign(__assign({}, prev), { SubCategory_key: subCategoryKey_1 })); });
        }
    };
    function handleInputChange(field, value) {
        setFormData(function (prev) {
            var _a;
            return (__assign(__assign({}, prev), (_a = {}, _a[field] = value, _a)));
        });
    }
    function validateForm() {
        var newErrors = {};
        if (!formData.requestedFor.trim())
            newErrors.requestedFor = isAr ? "متلقي الخدمة مطلوب" : "RequestedFor is required";
        if (!formData.Category.trim())
            newErrors.Category = isAr ? "الفئة مطلوبة" : "Category is required";
        if (!formData.SubCategory.trim())
            newErrors.SubCategory = isAr ? "الفئة الفرعية مطلوبة" : "SubCategory is required";
        if (!formData.Department.trim())
            newErrors.Department = isAr ? "القسم مطلوب" : "Department is required";
        if (!formData.Description.trim())
            newErrors.Description = isAr ? "الوصف مطلوب" : "Description is required";
        setErrors(newErrors);
        if (Object.keys(newErrors).length > 0) {
            props.onErrorRequiredFields();
            return false;
        }
        return true;
    }
    function handleSubmit() {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        setErrors({});
                        if (!validateForm())
                            return [2 /*return*/];
                        setShowSubmitbtn(true);
                        return [4 /*yield*/, props.onSave(formData)];
                    case 1:
                        _a.sent();
                        setShowSubmitbtn(false);
                        return [2 /*return*/];
                }
            });
        });
    }
    var displayName = (_b = props.userprofileAD) === null || _b === void 0 ? void 0 : _b.displayName;
    var initials = "";
    if (displayName && displayName.trim()) {
        var parts = displayName.split(" ");
        initials = parts[0][0] + parts[parts.length - 1][0];
    }
    var peoplePickerContext = {
        absoluteUrl: props.context.pageContext.web.absoluteUrl,
        msGraphClientFactory: props.context.msGraphClientFactory,
        spHttpClient: props.context.spHttpClient,
    };
    var _getPeoplePickerItems = function (selectedUserProfiles, internalName, internalName_text, internalName_key) { return __awaiter(void 0, void 0, void 0, function () {
        var emails, title, updatedErrors;
        return __generator(this, function (_a) {
            if (selectedUserProfiles.length > 0) {
                emails = selectedUserProfiles[0].id.split("|")[2];
                title = selectedUserProfiles[0].text;
                updatedErrors = __assign({}, errors);
                if (title.trim()) {
                    delete updatedErrors[internalName];
                }
                setErrors(updatedErrors);
                handleInputChange(internalName, emails);
                handleInputChange(internalName_text, title);
                _getUserRecId(emails, internalName_key);
            }
            else {
                handleInputChange(internalName, "");
                handleInputChange(internalName_text, "");
                handleInputChange(internalName_key, "");
            }
            return [2 /*return*/];
        });
    }); };
    var _getUserRecId = function (email, columnkey) { return __awaiter(void 0, void 0, void 0, function () {
        var response, rawResponse, jsonStart, jsonString, parsedData, RecId, error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!email)
                        return [2 /*return*/];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 5, , 6]);
                    return [4 /*yield*/, fetch(props.UserRecIdApilink, {
                            method: "GET",
                            headers: {
                                "Ocp-Apim-Subscription-Key": props.OcpApimKey,
                                Authorization: props.Authorization,
                                Email: email,
                            },
                        })];
                case 2:
                    response = _a.sent();
                    if (!response.ok) return [3 /*break*/, 4];
                    return [4 /*yield*/, response.text()];
                case 3:
                    rawResponse = _a.sent();
                    jsonStart = rawResponse.indexOf("{");
                    if (jsonStart !== -1) {
                        jsonString = rawResponse.slice(jsonStart);
                        parsedData = JSON.parse(jsonString);
                        RecId = parsedData.value[0].RecId;
                        handleInputChange(columnkey, RecId);
                    }
                    _a.label = 4;
                case 4: return [3 /*break*/, 6];
                case 5:
                    error_2 = _a.sent();
                    console.error("Error getting UserRecId:", error_2);
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    }); };
    var updateFormData = function (event, newValue, column) {
        var value = newValue !== null && newValue !== void 0 ? newValue : "";
        setFormData(function (prev) {
            var _a;
            return (__assign(__assign({}, prev), (_a = {}, _a[column] = value, _a)));
        });
        setErrors(function (prevErrors) {
            var newErrors = __assign({}, prevErrors);
            if (newErrors[column] && value.trim() !== "") {
                delete newErrors[column];
            }
            return newErrors;
        });
        forceUpdate();
    };
    var updateFormDropData = function (option, column) {
        setFormData(function (prev) {
            var _a;
            return (__assign(__assign({}, prev), (_a = {}, _a[column] = option === null || option === void 0 ? void 0 : option.key, _a)));
        });
        setErrors(function (prevErrors) {
            var newErrors = __assign({}, prevErrors);
            if (newErrors[column] && option.key) {
                delete newErrors[column];
            }
            return newErrors;
        });
        forceUpdate();
    };
    var forceUpdate = function () { return setForceUpdater(function (prev) { return prev + 1; }); };
    var removeAttachment = function (fileName) {
        var updatedFile = uploadedFiles.filter(function (file) { return file.name !== fileName; });
        if (fileInputRef.current) {
            fileInputRef.current.value = "";
        }
        setUploadedFiles(updatedFile);
        handleInputChange("files", updatedFile);
    };
    var getRandomNumber = function (min, max) {
        if (min === void 0) { min = 1; }
        if (max === void 0) { max = 1000; }
        return Math.floor(Math.random() * (max - min + 1)) + min;
    };
    function updateFileName(file, newName) {
        return new File([file], newName, {
            type: file.type,
            lastModified: file.lastModified,
        });
    }
    var randomNumber = getRandomNumber(1, 1000);
    var readFile = function (e, field) {
        var requesterFileList = e.target.files;
        if (requesterFileList && requesterFileList.length > 0) {
            var file_1 = requesterFileList[0];
            var reader_1 = new FileReader();
            reader_1.onload = function () {
                var fileContent = reader_1.result;
                if (fileContent === "") {
                    setShowEmptyFileModal(true);
                    e.target.value = "";
                }
                else {
                    setShowEmptyFileModal(false);
                    var fileExtension = file_1.name.substring(file_1.name.lastIndexOf(".") + 1, file_1.name.length);
                    var fileName = file_1.name
                        .substring(0, file_1.name.lastIndexOf(".") + 1)
                        .replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, "") +
                        randomNumber +
                        "." +
                        fileExtension;
                    var updatedFile = updateFileName(file_1, fileName);
                    var newFile_1 = {
                        name: fileName,
                        content: updatedFile,
                        progress: 0
                    };
                    setUploadedFiles(function (prevFiles) {
                        var updatedFiles = __spreadArray(__spreadArray([], prevFiles, true), [newFile_1], false);
                        setFormData(function (prev) {
                            var _a;
                            return (__assign(__assign({}, prev), (_a = {}, _a[field] = updatedFiles, _a)));
                        });
                        return updatedFiles;
                    });
                    var currentProgress_1 = 0;
                    var interval_1 = setInterval(function () {
                        if (currentProgress_1 >= 100) {
                            clearInterval(interval_1);
                        }
                        else {
                            currentProgress_1 += 10;
                            setUploadedFiles(function (prevFiles) {
                                return prevFiles.map(function (f) {
                                    return f.name === newFile_1.name ? __assign(__assign({}, f), { progress: currentProgress_1 }) : f;
                                });
                            });
                        }
                    }, 300);
                }
            };
            reader_1.onerror = function (error) {
                alert("Error reading the file: " + error);
            };
            reader_1.readAsText(file_1);
        }
    };
    return (React.createElement("div", null,
        showEmptyFileModal && (React.createElement(AlertModal, { showModal: true, handleShowModal: function () { return setShowEmptyFileModal(true); }, handleCloseModal: function () { return setShowEmptyFileModal(false); }, heading: "Warning", message: "This attachment has no content", style: "", section: "rejected", icon: "WarningSolid", isAr: isAr })),
        React.createElement("div", { className: "maincontainer" },
            React.createElement("div", { className: "header-top" },
                React.createElement("div", { className: "person-image" }, initials),
                React.createElement("div", null,
                    React.createElement("div", { className: "person-name" }, (_c = props.userprofileAD) === null || _c === void 0 ? void 0 : _c.displayName),
                    React.createElement("div", { className: "person-description" }, (_d = props.userprofileAD) === null || _d === void 0 ? void 0 :
                        _d.jobTitle,
                        " | ID: ",
                        props.EmpId ? props.EmpId : "N/A"))),
            React.createElement("div", { className: "textContainer" },
                React.createElement("h2", { className: "form-heading" }, isAr
                    ? "يرجى تعبئة جميع الحقول المطلوبة لتسهيل اعتماد الطلب"
                    : "Kindly complete the form below"),
                React.createElement("div", { className: "fieldContainer" },
                    React.createElement(TextField, { type: "text", label: isAr ? "تم التقديم بواسطة" : "Requested By", className: "form-text", readOnly: true, value: ((_e = props.userprofileAD) === null || _e === void 0 ? void 0 : _e.displayName) || "" }),
                    React.createElement("div", { className: "people-picker-wrapper ".concat(errors.requestedFor ? "error-border" : "") },
                        React.createElement(PeoplePicker, { context: peoplePickerContext, titleText: isAr ? "مطلوب لـ *" : "Requested For *", personSelectionLimit: 1, groupName: "", groupId: "", defaultSelectedUsers: [formData.requestedFor], showtooltip: true, disabled: false, searchTextLimit: 1, onChange: function (e) {
                                _getPeoplePickerItems(e, "requestedFor", "requestedFor_Title", "requestedFor_key");
                            }, principalTypes: [PrincipalType.User], resolveDelay: 200 })),
                    React.createElement(Dropdown, { label: isAr ? "الفئة *" : "Category *", selectedKey: formData.Category, className: "dropdownfield ".concat(errors.Category ? "error-field" : ""), styles: {
                            root: {
                                marginBottom: '0px'
                            },
                            dropdown: {
                                borderColor: errors.Category ? "red" : undefined,
                                height: '60px'
                            },
                            title: {
                                height: '60px',
                                display: 'flex',
                                alignItems: 'center',
                                padding: '0 12px',
                                borderRadius: '6px',
                                borderColor: '#ccc'
                            },
                            caretDownWrapper: {
                                height: '60px',
                                lineHeight: '60px'
                            }
                        }, onChange: handleCategoryChange, options: categoryOptions }),
                    React.createElement(Dropdown, { label: isAr ? "الفئة الفرعية *" : "SubCategory *", selectedKey: formData.SubCategory, disabled: !formData.Category, className: "dropdownfield ".concat(errors.SubCategory ? "error-field" : ""), styles: {
                            root: {
                                marginBottom: '0px'
                            },
                            dropdown: {
                                borderColor: errors.SubCategory ? "red" : undefined,
                                height: '60px'
                            },
                            title: {
                                height: '60px',
                                display: 'flex',
                                alignItems: 'center',
                                padding: '0 12px',
                                borderRadius: '6px',
                                borderColor: '#ccc'
                            },
                            caretDownWrapper: {
                                height: '60px',
                                lineHeight: '60px'
                            }
                        }, onChange: handleSubCategoryChange, options: subCategoryOptions }),
                    React.createElement(TextField, { label: isAr ? "القسم *" : "Department *", value: formData.Department, onChange: function (ev, newValue) { return updateFormData(ev, newValue, "Department"); }, className: "form-text ".concat(errors.Department ? "error-field" : "") }),
                    React.createElement(TextField, { label: isAr ? "الموقع" : "Location", value: formData.Location, onChange: function (ev, newValue) { return updateFormData(ev, newValue, "Location"); }, className: "form-text" })),
                React.createElement("div", { className: "description_div" },
                    React.createElement(TextField, { label: isAr ? "الوصف *" : "Description *", value: formData.Description, multiline: true, rows: 4, type: "text-area", className: "text-area ".concat(errors.Description ? "error-field" : ""), onChange: function (ev, newValue) { return updateFormData(ev, newValue, "Description"); }, styles: {
                            root: { color: "#555" },
                            fieldGroup: { border: "1px solid #ccc", borderColor: errors.Description ? "red" : undefined },
                            field: { color: "#555" },
                        } })),
                React.createElement(Col, { className: "mt-4" },
                    React.createElement("div", { style: { display: "flex", alignItems: "end" } },
                        React.createElement("label", { className: "attach_label", style: {
                                marginRight: "4px",
                                marginTop: "24px",
                                fontSize: "14px",
                                fontFamily: "Segoe UI",
                                color: "#555",
                                fontWeight: "600",
                                marginBottom: "11.5px",
                            } }, isAr
                            ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                            : "Please attach any supporting documents or pictures (optional).")),
                    React.createElement("div", { className: "attachment-container" },
                        React.createElement("div", { className: "attachment-placeholder" },
                            React.createElement("img", { className: "attachment-icon", src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKoSURBVHgB7ZnbjdpAFIaPDYLX7WDdQUBcJJ4CFWRTQUgHbAVABbAVQCpIOljyhIS4leAO4rwB4pL/sMcbxxdij8dWFO0vjcY7c8Z8mjOeOXPWoAgtFotKsVgc4bGCckfpZO/3+06r1bIpoYywxuVy+WCa5kTAHJSfpK57qZUgA4Cbzca6XC7PeOT6CaCDarXqkKLW67WdBtL0N5zP567ATWu1Wi8NnE9bfm+5XH7GJMReMgFAwzDeXztM8wtpFN7bIYFkD8WFNKM6drudTRrFnhBIG6USF9KkHBUC+fVvY3IFZAHS9kC2V6vV5JZ9boDz+dxyn72QqLu3IDMHhBu/cV0qlfre9hDIUdj4ImUs7AZjQH5gCOyJbXpx7VVof7VDfw+QDra24R/jKWN5Zuo7ioXS9hWLfkN2/eMzn0EWQzKMbCuVMBs5vQLSCggX9VDdw02PYf1yKs3C+uD+sGZ9LgZcHy4ayVqakCZpARS4gfu3fJV90qDUgD44N7DgE2OgAzIVIOLGTy4cFvlnkrjxeDx2dEEqA0pQOxW4R3wYU7ev2WxuT6fTR34WyB4pSgmQrwMScTPcEHBjv02j0ZjJrDLkiGebFKQEWCgUHlDdCdwgyo5nlWeXXiCrpCAlQIbiH7wF57Ed80nCVwdSkPJGjU13m8B2RorKPR5MqjfAtPrnAXWHWz9QLqRRWgElMNUqrYAasxCv+r/WoBz67yiFcPRt6/X6U1z7RIBYY3zgVyiF8A4bVWaAHEJZlE6xj0hWIkC5ndmUo8I+kmtUjPSvRTkJ11F32QR2gQAgFvGMa8R8Wi49cYTf7EkdcH8AUMJ4myTzxNEzZSROKElOhj8+53A4DP02oUl0b56a8pHDFy2+y/g7jKgRAjkgPf+GiBRnvzBz46jE+i8JiDR7F2tlUAAAAABJRU5ErkJggg==", alt: "Attachment Icon" }),
                            isAr
                                ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                                : "Please attach any supporting documents or pictures (optional)",
                            React.createElement("input", { type: "file", ref: fileInputRef, multiple: true, onChange: function (e) {
                                    readFile(e, "files");
                                } })),
                        React.createElement("span", { style: { color: "red" } }, errors.files || showErrorUpload)),
                    uploadedFiles.map(function (file, index) { return (React.createElement("div", { key: index },
                        React.createElement("div", { className: "uploadeditems" },
                            React.createElement("strong", null, file.name),
                            React.createElement("div", { className: "progresscontainer" },
                                React.createElement("div", { className: "progressbar", id: "progressbar", style: { width: "".concat(file.progress || 0, "%") } })),
                            React.createElement("div", { className: "cancelbtn", onClick: function () {
                                    removeAttachment(file.name);
                                } }, "X")))); })),
                React.createElement("div", { className: "buttonContainer" },
                    React.createElement(PrimaryButton, { disabled: showSubmitbtn, onClick: function () {
                            handleSubmit();
                        }, styles: { root: { fontSize: "20px" } }, text: isAr ? "تقديم" : "Submit", className: "submit-formbtn" }),
                    React.createElement(DefaultButton, { text: isAr ? "مسح" : "Clear", className: "cancel-formbtn", onClick: function () {
                            var _a;
                            setFormData({
                                requestedBy: (_a = props.userprofileAD) === null || _a === void 0 ? void 0 : _a.displayName,
                                requestedFor: "",
                                requestedFor_Title: "",
                                requestedFor_key: "",
                                Category: "",
                                Category_key: "",
                                SubCategory: "",
                                SubCategory_key: "",
                                Department: "",
                                Location: "",
                                Description: "",
                                files: [],
                            });
                            setUploadedFiles([]);
                            setShowErrorUpload("");
                            setErrors({});
                            if (inputRef.current)
                                inputRef.current.value = "";
                        } }))))));
};
export default MaintenanceRequestUIForm;
//# sourceMappingURL=MaintenanceRequestUIForm.js.map