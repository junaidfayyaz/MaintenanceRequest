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
      let payload = {
        [props.ProfileLink_col]: formData.requestedBy_key,
        [props.Category_col]: formData.Category,
        [props.SubCategory_col]: formData.SubCategory,
        [props.Department_col]: formData.Department,
        [props.Location_col]: formData.Location,
        [props.Description_col]: formData.Description,
        [props.SubmittedBy_col]: formData.requestedBy,
        "Service": "Report An Issue",
        "Team": "Service Desk",
        "OwnerTeam": "Service Desk",
        "Status": "Logged",
        "Subject": formData.Category ? `Maintenance Request - ${formData.Category}` : "Maintenance Request",
        "Summary": formData.Category ? `Maintenance Request - ${formData.Category}` : "Maintenance Request",
        "Symptom": formData.Description ? formData.Description : "Maintenance Request",
        "Description": formData.Description ? formData.Description : "Maintenance Request",
        "Impact": "Low",
        "Urgency": "Low"
      };

      console.log("=== VERIFY DATA ===");
      console.log("Raw Form Data:", formData);
      console.log("API Payload:", payload);
      console.log("===================");

      const response = await fetch(props.Apilink, {
        method: "POST",
        headers: {
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

      const strRequestNum = parsedData?.IncidentNumber || parsedData?.ServiceReqNumber;
      const requestRecId = parsedData?.RecId;

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
      ApiformData.append("ObjectType", "Incident#"); // Based on API, might need to change if using ServiceReq
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
