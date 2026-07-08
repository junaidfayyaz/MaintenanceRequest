import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../css/style.css";
import {
  IUserProfile,
  IMaintenanceRequestFormData,
} from "../webparts/maintenanceRequest/components/IMaintenanceRequestProps";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Col } from "react-bootstrap";
import AlertModal from "./alertModal/AlertModal";
import { Web } from "@pnp/sp/webs";
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

const isAr =
  window.location.pathname.includes("/ar/") ||
  window.location.search.includes("lang=ar");

const MaintenanceRequestUIForm: React.FC<IRequestUIFormProps> = (props) => {
  const [formData, setFormData] = useState<IMaintenanceRequestFormData>({
    requestedBy: props.userprofileAD?.displayName,
    requestedFor: "",
    requestedFor_Title: "",
    requestedFor_key: "",
    Category: "",
    SubCategory: "",
    Department: "",
    Location: "",
    Description: "",
    files: [],
  });

  const [uploadedFiles, setUploadedFiles] = useState<Array<{ name: string; content?: any; progress?: number }>>([]);
  const [showErrorUpload, setShowErrorUpload] = useState("");
  const [showSubmitbtn, setShowSubmitbtn] = useState(false);
  const [errors, setErrors] = useState<{ [field: string]: string }>({});
  const inputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [showEmptyFileModal, setShowEmptyFileModal] = useState(false);
  const [, setForceUpdater] = useState(0);

  const [categoryOptions, setCategoryOptions] = useState<IDropdownOption[]>([]);
  const [subCategoryOptions, setSubCategoryOptions] = useState<IDropdownOption[]>([]);
  const [listData, setListData] = useState<any[]>([]);

  useEffect(() => {
    _getUserRecId(props.CurrentUserEmail, "requestedBy_key");
    fetchCategories();
  }, []);

  const fetchCategories = async () => {
    try {
      // Access specific site for the list
      const tahyahWeb = Web("https://mcitgovqa.sharepoint.com/sites/Tahyah");
      const items = await tahyahWeb.lists
        .getByTitle("MaintenanceRequestCategoryList")
        .items.getAll();
      
      setListData(items);

      // Extract unique categories
      const uniqueCategories = Array.from(new Set(items.map((item) => item.Category).filter(Boolean)));
      
      setCategoryOptions(
        uniqueCategories.map((cat) => ({ key: cat as string, text: cat as string }))
      );
    } catch (error) {
      console.error("Error fetching categories: ", error);
    }
  };

  const handleCategoryChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const selectedCategory = option.key as string;
      updateFormDropData(option, "Category");

      const categoryItem = listData.find((item) => item.Category === selectedCategory);
      let categoryKey = "";
      if (categoryItem) {
        // Look for any property name that might contain the Category recId
        const possibleKeys = Object.keys(categoryItem).filter(k => k.toLowerCase().includes("recid") && !k.toLowerCase().includes("sub"));
        if (possibleKeys.length > 0) {
          categoryKey = categoryItem[possibleKeys[0]];
        }
      }
      
      // Fallback to user-provided mapping if not found in SharePoint list
      if (!categoryKey && selectedCategory === "Access Control") {
        categoryKey = "87967EC7EA98409C8C2A636924C6C797";
      }

      setFormData((prev) => ({ 
        ...prev, 
        Category_key: categoryKey,
        SubCategory: "",
        SubCategory_key: ""
      }));

      // Filter subcategories based on selected category and ensure uniqueness
      const filteredSubCats = Array.from(
        new Set(
          listData
            .filter((item) => item.Category === selectedCategory)
            .map((item) => item.SubCategory)
            .filter(Boolean)
        )
      );

      setSubCategoryOptions(
        filteredSubCats.map((subcat) => ({ key: subcat as string, text: subcat as string }))
      );
    }
  };

  const handleSubCategoryChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const selectedSubCategory = option.key as string;
      updateFormDropData(option, "SubCategory");
      
      const subCategoryItem = listData.find((item) => item.Category === formData.Category && item.SubCategory === selectedSubCategory);
      let subCategoryKey = "";
      if (subCategoryItem) {
        // Look for any property name that might contain the SubCategory recId
        const possibleKeys = Object.keys(subCategoryItem).filter(k => k.toLowerCase().includes("sub") && k.toLowerCase().includes("recid"));
        if (possibleKeys.length > 0) {
          subCategoryKey = subCategoryItem[possibleKeys[0]];
        } else {
           // fallback if it just named recId on the item level
           const fallbackKeys = Object.keys(subCategoryItem).filter(k => k.toLowerCase().includes("recid"));
           if (fallbackKeys.length > 0) {
             // pick one that is not the category key
             const subKey = fallbackKeys.find(k => subCategoryItem[k] !== formData.Category_key);
             if (subKey) subCategoryKey = subCategoryItem[subKey];
           }
        }
      }

      // Fallback to user-provided mapping if not found in SharePoint list
      if (!subCategoryKey && selectedSubCategory === "automated door noisy") {
        subCategoryKey = "6BFF811DE98A4C17B21D38E084DEB1BB";
      }

      setFormData((prev) => ({ ...prev, SubCategory_key: subCategoryKey }));
    }
  };

  function handleInputChange(field: string, value: any) {
    setFormData((prev) => ({ ...prev, [field]: value }));
  }

  function validateForm() {
    const newErrors: { [field: string]: string } = {};

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

  async function handleSubmit() {
    setErrors({});
    if (!validateForm()) return;
    setShowSubmitbtn(true);
    await props.onSave(formData);
    setShowSubmitbtn(false);
  }

  const displayName = props.userprofileAD?.displayName;
  let initials = "";
  if (displayName && displayName.trim()) {
    const parts = displayName.split(" ");
    initials = parts[0][0] + parts[parts.length - 1][0];
  }

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory as any,
    spHttpClient: props.context.spHttpClient as any,
  };

  const _getPeoplePickerItems = async (
    selectedUserProfiles: any[],
    internalName: string,
    internalName_text: string,
    internalName_key: string
  ) => {
    if (selectedUserProfiles.length > 0) {
      const emails = selectedUserProfiles[0].id.split("|")[2];
      const title = selectedUserProfiles[0].text;
      const updatedErrors = { ...errors };
      if (title.trim()) {
        delete updatedErrors[internalName];
      }
      setErrors(updatedErrors);
      handleInputChange(internalName, emails);
      handleInputChange(internalName_text, title);
      _getUserRecId(emails, internalName_key);
    } else {
      handleInputChange(internalName, "");
      handleInputChange(internalName_text, "");
      handleInputChange(internalName_key, "");
    }
  };

  const _getUserRecId = async (email: string | undefined, columnkey: string) => {
    if (!email) return;
    try {
      const response = await fetch(props.UserRecIdApilink, {
        method: "GET",
        headers: {
          "Ocp-Apim-Subscription-Key": props.OcpApimKey,
          Authorization: props.Authorization,
          Email: email,
        },
      });
      if (response.ok) {
        const rawResponse = await response.text();
        const jsonStart = rawResponse.indexOf("{");
        if (jsonStart !== -1) {
          const jsonString = rawResponse.slice(jsonStart);
          const parsedData = JSON.parse(jsonString);
          let RecId = parsedData.value[0].RecId;
          handleInputChange(columnkey, RecId);
        }
      }
    } catch (error: any) {
      console.error("Error getting UserRecId:", error);
    }
  };

  const updateFormData = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string | undefined,
    column: any
  ) => {
    const value = newValue ?? "";
    setFormData((prev) => ({
      ...prev,
      [column]: value,
    }));

    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && value.trim() !== "") {
        delete newErrors[column];
      }
      return newErrors;
    });

    forceUpdate();
  };

  const updateFormDropData = (option: any, column: any) => {
    setFormData((prev) => ({ ...prev, [column]: option?.key as string }));
    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && option.key) {
        delete newErrors[column];
      }
      return newErrors;
    });
    forceUpdate();
  };

  const forceUpdate = () => setForceUpdater((prev) => prev + 1);

  const removeAttachment = (fileName: string) => {
    const updatedFile = uploadedFiles.filter((file) => file.name !== fileName);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setUploadedFiles(updatedFile);
    handleInputChange("files", updatedFile);
  };

  const getRandomNumber = (min: number = 1, max: number = 1000): number => {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  };

  function updateFileName(file: File, newName: string) {
    return new File([file], newName, {
      type: file.type,
      lastModified: file.lastModified,
    });
  }

  const randomNumber = getRandomNumber(1, 1000);

  const readFile = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    const requesterFileList = e.target.files;
    if (requesterFileList && requesterFileList.length > 0) {
      const file = requesterFileList[0];
      const reader = new FileReader();
      reader.onload = () => {
        const fileContent = reader.result;
        if (fileContent === "") {
          setShowEmptyFileModal(true);
          e.target.value = "";
        } else {
          setShowEmptyFileModal(false);
          const fileExtension = file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
          const fileName =
            file.name
              .substring(0, file.name.lastIndexOf(".") + 1)
              .replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, "") +
            randomNumber +
            "." +
            fileExtension;
          
          const updatedFile = updateFileName(file as File, fileName);
          const newFile = {
            name: fileName,
            content: updatedFile,
            progress: 0
          };
          
          setUploadedFiles((prevFiles) => {
            const updatedFiles = [...prevFiles, newFile];
            setFormData((prev) => ({ ...prev, [field]: updatedFiles }));
            return updatedFiles;
          });
          
          let currentProgress = 0;
          const interval = setInterval(() => {
            if (currentProgress >= 100) {
              clearInterval(interval);
            } else {
              currentProgress += 10;
              setUploadedFiles((prevFiles) =>
                prevFiles.map((f) =>
                  f.name === newFile.name ? { ...f, progress: currentProgress } : f
                )
              );
            }
          }, 300);
        }
      };

      reader.onerror = (error) => {
        alert("Error reading the file: " + error);
      };
      reader.readAsText(file as Blob);
    }
  };

  return (
    <div>
      {showEmptyFileModal && (
        <AlertModal
          showModal={true}
          handleShowModal={() => setShowEmptyFileModal(true)}
          handleCloseModal={() => setShowEmptyFileModal(false)}
          heading={"Warning"}
          message={"This attachment has no content"}
          style={""}
          section={"rejected"}
          icon={"WarningSolid"}
          isAr={isAr}
        />
      )}
      <div className="maincontainer">
        <div className="header-top">
          <div className="person-image">{initials}</div>
          <div>
            <div className="person-name">{props.userprofileAD?.displayName}</div>
            <div className="person-description">
              {props.userprofileAD?.jobTitle} | ID: {props.EmpId ? props.EmpId : "N/A"}
            </div>
          </div>
        </div>
        <div className="textContainer">
          <h2 className="form-heading">
            {isAr
              ? "يرجى تعبئة جميع الحقول المطلوبة لتسهيل اعتماد الطلب"
              : "Kindly complete the form below"}
          </h2>

          <div className="fieldContainer">
            {/* Requested By */}
            <TextField
              type="text"
              label={isAr ? "تم التقديم بواسطة" : "Requested By"}
              className="form-text"
              readOnly
              value={props.userprofileAD?.displayName || ""}
            />
            
            {/* Requested For */}
            <div className={`people-picker-wrapper ${errors.requestedFor ? "error-border" : ""}`}>
              <PeoplePicker
                context={peoplePickerContext}
                titleText={isAr ? "مطلوب لـ *" : "Requested For *"}
                personSelectionLimit={1}
                groupName={""}
                groupId=""
                defaultSelectedUsers={[formData.requestedFor]}
                showtooltip={true}
                disabled={false}
                searchTextLimit={1}
                onChange={(e) => {
                  _getPeoplePickerItems(e, "requestedFor", "requestedFor_Title", "requestedFor_key");
                }}
                principalTypes={[PrincipalType.User]}
                resolveDelay={200}
              />
            </div>

            {/* Category Dropdown */}
            <Dropdown
              label={isAr ? "الفئة *" : "Category *"}
              selectedKey={formData.Category}
              className={`dropdownfield ${errors.Category ? "error-field" : ""}`}
              styles={{
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
              }}
              onChange={handleCategoryChange}
              options={categoryOptions}
            />

            {/* SubCategory Dropdown */}
            <Dropdown
              label={isAr ? "الفئة الفرعية *" : "SubCategory *"}
              selectedKey={formData.SubCategory}
              disabled={!formData.Category}
              className={`dropdownfield ${errors.SubCategory ? "error-field" : ""}`}
              styles={{
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
              }}
              onChange={handleSubCategoryChange}
              options={subCategoryOptions}
            />

            {/* Department Textbox */}
            <TextField
              label={isAr ? "القسم *" : "Department *"}
              value={formData.Department}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Department")}
              className={`form-text ${errors.Department ? "error-field" : ""}`}
            />

            {/* Location Textbox (Optional) */}
            <TextField
              label={isAr ? "الموقع" : "Location"}
              value={formData.Location}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Location")}
              className="form-text"
            />
          </div>

          <div className="description_div">
            {/* Description Multi-line */}
            <TextField
              label={isAr ? "الوصف *" : "Description *"}
              value={formData.Description}
              multiline
              rows={4}
              type="text-area"
              className={`text-area ${errors.Description ? "error-field" : ""}`}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Description")}
              styles={{
                root: { color: "#555" },
                fieldGroup: { border: "1px solid #ccc", borderColor: errors.Description ? "red" : undefined },
                field: { color: "#555" },
              }}
            />
          </div>

          <Col className="mt-4">
            <div style={{ display: "flex", alignItems: "end" }}>
              <label
                className="attach_label"
                style={{
                  marginRight: "4px",
                  marginTop: "24px",
                  fontSize: "14px",
                  fontFamily: "Segoe UI",
                  color: "#555",
                  fontWeight: "600",
                  marginBottom: "11.5px",
                }}
              >
                {isAr
                  ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                  : "Please attach any supporting documents or pictures (optional)."}
              </label>
            </div>

            <div className="attachment-container">
              <div className="attachment-placeholder">
                <img
                  className="attachment-icon"
                  src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKoSURBVHgB7ZnbjdpAFIaPDYLX7WDdQUBcJJ4CFWRTQUgHbAVABbAVQCpIOljyhIS4leAO4rwB4pL/sMcbxxdij8dWFO0vjcY7c8Z8mjOeOXPWoAgtFotKsVgc4bGCckfpZO/3+06r1bIpoYywxuVy+WCa5kTAHJSfpK57qZUgA4Cbzca6XC7PeOT6CaCDarXqkKLW67WdBtL0N5zP567ATWu1Wi8NnE9bfm+5XH7GJMReMgFAwzDeXztM8wtpFN7bIYFkD8WFNKM6drudTRrFnhBIG6USF9KkHBUC+fVvY3IFZAHS9kC2V6vV5JZ9boDz+dxyn72QqLu3IDMHhBu/cV0qlfre9hDIUdj4ImUs7AZjQH5gCOyJbXpx7VVof7VDfw+QDra24R/jKWN5Zuo7ioXS9hWLfkN2/eMzn0EWQzKMbCuVMBs5vQLSCggX9VDdw02PYf1yKs3C+uD+sGZ9LgZcHy4ayVqakCZpARS4gfu3fJV90qDUgD44N7DgE2OgAzIVIOLGTy4cFvlnkrjxeDx2dEEqA0pQOxW4R3wYU7ev2WxuT6fTR34WyB4pSgmQrwMScTPcEHBjv02j0ZjJrDLkiGebFKQEWCgUHlDdCdwgyo5nlWeXXiCrpCAlQIbiH7wF57Ed80nCVwdSkPJGjU13m8B2RorKPR5MqjfAtPrnAXWHWz9QLqRRWgElMNUqrYAasxCv+r/WoBz67yiFcPRt6/X6U1z7RIBYY3zgVyiF8A4bVWaAHEJZlE6xj0hWIkC5ndmUo8I+kmtUjPSvRTkJ11F32QR2gQAgFvGMa8R8Wi49cYTf7EkdcH8AUMJ4myTzxNEzZSROKElOhj8+53A4DP02oUl0b56a8pHDFy2+y/g7jKgRAjkgPf+GiBRnvzBz46jE+i8JiDR7F2tlUAAAAABJRU5ErkJggg=="
                  alt="Attachment Icon"
                />
                {isAr
                  ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                  : "Please attach any supporting documents or pictures (optional)"}

                <input
                  type="file"
                  ref={fileInputRef}
                  multiple={true}
                  onChange={(e) => {
                    readFile(e, "files");
                  }}
                />
              </div>
              <span style={{ color: "red" }}>
                {errors.files || showErrorUpload}
              </span>
            </div>

            {uploadedFiles.map((file, index) => (
              <div key={index}>
                <div className="uploadeditems">
                  <strong>{file.name}</strong>
                  <div className="progresscontainer">
                    <div
                      className="progressbar"
                      id="progressbar"
                      style={{ width: `${file.progress || 0}%` }}
                    ></div>
                  </div>
                  <div
                    className="cancelbtn"
                    onClick={() => {
                      removeAttachment(file.name);
                    }}
                  >
                    X
                  </div>
                </div>
              </div>
            ))}
          </Col>

          <div className="buttonContainer">
            <PrimaryButton
              disabled={showSubmitbtn}
              onClick={() => {
                handleSubmit();
              }}
              styles={{ root: { fontSize: "20px" } }}
              text={isAr ? "تقديم" : "Submit"}
              className="submit-formbtn"
            />
            <DefaultButton
              text={isAr ? "مسح" : "Clear"}
              className="cancel-formbtn"
              onClick={() => {
                setFormData({
                  requestedBy: props.userprofileAD?.displayName,
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
                if (inputRef.current) inputRef.current.value = "";
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default MaintenanceRequestUIForm;
