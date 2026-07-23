import * as React from "react";
import { useEffect, useState } from "react";
import { MSGraphClient } from "@microsoft/sp-http";
import styles from "./MaintenanceRequest.module.scss";
import type { IMaintenanceRequestProps, IMaintenanceRequestFormData, IUserProfile } from "./IMaintenanceRequestProps";
import { escape } from "@microsoft/sp-lodash-subset";
import AlertModal from "../../../components/alertModal/AlertModal";
import MaintenanceRequestUIForm from "../../../components/MaintenanceRequestUIForm";

const isAr =
  window.location.pathname.includes("/ar/") ||
  window.location.search.includes("lang=ar");

const MaintenanceRequest: React.FC<IMaintenanceRequestProps> = (props) => {
  const [userProfileAD, setUserProfileAD] = useState<IUserProfile | null>(null);
  const [isLoadingUser, setIsLoadingUser] = useState<boolean>(true);
  const [showModal, setShowModal] = useState(false);
  const [modalHeading, setModalHeading] = useState("");
  const [modalMessage, setModalMessage] = useState("");
  const [alertsection, setAlertsection] = useState("");
  const [iconLoad, setIconLoad] = useState("");
  
  const handleShowModal = () => setShowModal(true);
  
  const handleCloseModal1 = (section: string) => {
    if (section === "Accepted") {
      setShowModal(false);
      window.location.reload();
    } else {
      setShowModal(false);
    }
  };

  useEffect(() => {
    (async () => {
      try {
        const client: MSGraphClient =
          await props.context.msGraphClientFactory.getClient("3");
        const userAD: any = await client
          .api("/me")
          .select(
            "displayName,jobTitle,department,employeeId,mail,onPremisesExtensionAttributes"
          )
          .get();

        const userProfile: IUserProfile = {
          displayName: userAD.displayName || "",
          jobTitle: userAD.jobTitle || "",
          department: userAD.department || "",
          mail: userAD.mail || "",
          employeeId:
            userAD?.onPremisesExtensionAttributes?.extensionAttribute15 || "",
        };
        setUserProfileAD(userProfile);
        setIsLoadingUser(false);
      } catch (error) {
        console.error("Error fetching user info:", error);
        setIsLoadingUser(false);
      }
    })();
  }, [props]);

  const showErrorModal = () => {
    setModalHeading("Warning");
    setModalMessage("Please fill Required fields");
    setAlertsection("rejected");
    setIconLoad("WarningSolid");
    handleShowModal();
  };

  const saveRequest = async (formData: IMaintenanceRequestFormData) => {
    try {
      let subCatCol = props.SubCategory_col;
      if (formData.Category === "Health") {
        subCatCol = "par-98AFADB1E55E4A20A13E33E1FB3AAF38";
      } else if (formData.Category === "Air Conditioning / Ventilation") {
        subCatCol = "par-B035253478874BE79920984A1F401FFB";
      } else if (formData.Category === "Electrical / Lighting") {
        subCatCol = "par-D2D07B3A41DD4F16AD18A113344F6D86";
      } else if (formData.Category === "Plumbing & Water Systems") {
        subCatCol = "par-7CB209DA0FB045BE9CB5DFC4EFFB5A1B";
      } else if (formData.Category === "Interior Fit-out") {
        subCatCol = "par-DE49E6FD7AE5424A8DC1FDFAC56F6C3E";
      } else if (formData.Category === "Fire Systems") {
        subCatCol = "par-D08FBDD8FCD4434B85BEF12F75CF4B59";
      } else if (formData.Category === "Elevators") {
        subCatCol = "par-DCB8F64B2B274DF1852C51C0DC97147C";
      } else if (formData.Category === "Cleaning") {
        subCatCol = "par-E1D98ECA66DF4D4F98A625ACE4482589";
      } else if (formData.Category === "Security") {
        subCatCol = "par-2A2B853272424F679E129BA996561AAC";
      } else if (formData.Category === "Pest Control") {
        subCatCol = "par-8B027D154F64449BBC37690497DBB611";
      } else if (formData.Category === "Landscaping") {
        subCatCol = "par-242A02793B8B4D5BB8FEA98AFCAEA9F5";
      } else if (formData.Category === "Mail Room") {
        subCatCol = "par-6276B7DC33804760B17A9623EDB04806";
      } else if (formData.Category === "Pantry Service") {
        subCatCol = "par-35C4083864304869AD82C9042D97B697";
      } else if (formData.Category === "Transportation") {
        subCatCol = "par-1B9112137E2A4B028551FF05F3DCE00A";
      } else if (formData.Category === "Meeting Room") {
        subCatCol = "par-68701F56FB6C468A8C3C36D98213187B";
      } else if (formData.Category === "Environment") {
        subCatCol = "par-1EEE6900356A4924B3418EAA74B2CE0F";
      } else if (formData.Category === "WorkSpace") {
        subCatCol = "par-C8387403403C467B9CF1F6E3C33C1B8C";
      } else if (formData.Category === "Quality") {
        subCatCol = "par-4FCD70FCF19643D08CB6FDE4CB8D4200";
      } else if (formData.Category === "Others") {
        subCatCol = "par-C06F7C725B924E09ABF2D8EBBB103CA8";
      }

      const requestedForUsername = formData.requestedFor_Title ? formData.requestedFor_Title : (formData.requestedFor ? formData.requestedFor.split('@')[0] : "");

      let params: any = {
        "ProfileLink": formData.requestedFor_key,
        "Source": "Tahyya",
        [props.ProfileLink_col]: formData.requestedFor_key,
        [props.Category_col]: formData.Category,
        [subCatCol]: formData.SubCategory,
        [props.Department_col]: formData.Department,
        [props.Location_col]: formData.Location,
        [props.Description_col]: formData.Description,
        [props.SubmittedBy_col]: formData.requestedBy,
        "RequestedBy": formData.requestedBy,
        "RequestedFor": requestedForUsername,
        "par-1FAC66E535124E2AA853A597E2994435": formData.requestedBy,
        "par-1FAC66E535124E2AA853A597E2994435-recId": formData.requestedBy_key,
        "par-8ADEEF4D4DEA4949A303E34C6D135E71": requestedForUsername,
        "par-8ADEEF4D4DEA4949A303E34C6D135E71-recId": formData.requestedFor_key,
      };

      if (formData.Category_key) {
        params[`${props.Category_col}-recId`] = formData.Category_key;
      }
      if (formData.SubCategory_key) {
        params[`${subCatCol}-recId`] = formData.SubCategory_key;
      }

      let payload = {
        attachmentsToDelete: [],
        attachmentsToUpload: [],
        parameters: params,
        delayedFulfill: false,
        formName: "ServiceReq.ResponsiveAnalyst.DefaultLayout",
        saveReqState: false,
        serviceReqData: {
          Subject: formData.Category ? `Maintenance Request - ${formData.Category}` : "Maintenance Request",
          Symptom: formData.Description ? formData.Description : "Maintenance Request",
          Category: "",
          CreatedBy: formData.requestedBy,
          RequestedBy: formData.requestedBy,
          RequestedFor: requestedForUsername,
          Source: "Tahyya"
        },
        subscriptionId: "EE937A037A90490D8A38B2437B567673",
      };

      console.log("=== VERIFY DATA ===");
      console.log("Raw Form Data:", formData);
      console.log("API Payload:", payload);
      console.log("===================");

      const response = await fetch(props.Apilink, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          accept: "application/json;odata=verbose",
          "Ocp-Apim-Subscription-Key": props.OcpApimKey,
          Authorization: props.Authorization,
          Email: formData.requestedFor,
        },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Request failed: ${response.status} - ${errorText}`);
      }
      
      const rawResponse = await response.text();
      const jsonStart = rawResponse.indexOf("{");
      if (jsonStart === -1) {
        throw new Error("JSON not found in response");
      }

      const jsonString = rawResponse.slice(jsonStart);
      let parsedData;
      try {
        parsedData = JSON.parse(jsonString);
      } catch (e) {
        throw new Error("Failed to parse JSON: " + e.message);
      }

      const dataObj = Array.isArray(parsedData) ? parsedData[0] : parsedData;
      const rpcData = dataObj?.result?.ServiceRequests?.[0] || dataObj?.ServiceRequests?.[0];
      const strRequestNum = rpcData?.strRequestNum || dataObj?.IncidentNumber || dataObj?.ServiceReqNumber || dataObj?.requestNumber || parsedData?.strRequestNum || dataObj?.ServiceReqNum;
      const requestRecId = rpcData?.strRequestRecId || dataObj?.RecId || dataObj?.recId || parsedData?.RecId;

      console.log("=== API RESPONSE RESULT ===");
      console.log("Full Response Data:", JSON.stringify(parsedData, null, 2));
      console.log("Extracted ID (strRequestNum):", strRequestNum);
      console.log("Extracted RecId:", requestRecId);
      console.log("===========================");

      let flag = true;
      if (formData.files && formData.files.length > 0) {
        flag = false;
        formData.files.forEach(async (element: any) => {
          await saveRequestAttachment(
            requestRecId,
            strRequestNum,
            element,
            formData.requestedFor
          );
        });
      }

      if (flag) {
        setModalHeading(isAr ? "تمت العمليه بنجاح" : "Success");
        setModalMessage(
          isAr
            ? `تم تقديم طلبك بنجاح. رقم التعريف:${strRequestNum}`
            : `Your Request has been submitted successfully. ID:${strRequestNum}`
        );
        setAlertsection("Accepted");
        setIconLoad("SkypeCircleCheck");
        handleShowModal();
      }
      
    } catch (error: any) {
      console.error("Error submitting Request:", error);
      setModalHeading("Error");
      setModalMessage(error.message);
      setAlertsection("rejected");
      setIconLoad("ErrorBadge");
      handleShowModal();
    }
  };

  const saveRequestAttachment = async (
    recid: string,
    requestnum: string,
    fileObj: any,
    userEmail: string
  ) => {
    try {
      const ApiformData = new FormData();
      ApiformData.append("ObjectID", recid);
      ApiformData.append("ObjectType", "ServiceReq#");
      ApiformData.append("File", fileObj.content, fileObj.name);
      
      const apiUrl = props.attachmentApilink;
      
      const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Ocp-Apim-Subscription-Key": props.OcpApimKey,
          Authorization: props.Authorization,
          Email: userEmail,
        },
        body: ApiformData,
      });
      
      if (response.ok) {
        setModalHeading(isAr ? "تمت العمليه بنجاح" : "Success");
        setModalMessage(
          isAr
            ? `تم تقديم طلبك بنجاح. رقم التعريف:${requestnum}`
            : `Your Request has been submitted successfully. ID:${requestnum}`
        );
        setAlertsection("Accepted");
        setIconLoad("SkypeCircleCheck");
        handleShowModal();
      } else {
        const errorText = await response.text();
        throw new Error(`Request failed: ${response.status} - ${errorText}`);
      }
    } catch (error: any) {
      console.error("Error submitting Attachment:", error);
      setModalHeading("Error");
      setModalMessage(error.message);
      setAlertsection("rejected");
      setIconLoad("ErrorBadge");
      handleShowModal();
    }
  };

  if (isLoadingUser) {
    return <div>Loading user information...</div>;
  }

  return (
    <>
      <MaintenanceRequestUIForm
        Authorization={props.Authorization}
        OcpApimKey={props.OcpApimKey}
        UserRecIdApilink={props.UserRecIdApilink}
        context={props.context}
        userprofileAD={userProfileAD!}
        EmpId={userProfileAD?.employeeId || ""}
        CurrentUserEmail={props.CurrentUserEmail || ""}
        onErrorRequiredFields={() => showErrorModal()}
        onSave={async (formData) => {
          await saveRequest(formData);
        }}
      />

      <AlertModal
        showModal={showModal}
        handleShowModal={handleShowModal}
        handleCloseModal={handleCloseModal1}
        heading={modalHeading}
        message={modalMessage}
        style={""}
        section={alertsection}
        icon={iconLoad}
        isAr={isAr}
      />
    </>
  );
};

export default MaintenanceRequest;
