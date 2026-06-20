export interface IMaintenanceRequestProps {
  context: any;
  Apilink: string;
  attachmentApilink: string;
  UserRecIdApilink: string;
  OcpApimKey: string;
  Authorization: string;
  CurrentUserEmail?: string;

  Category_col: string;
  SubCategory_col: string;
  Department_col: string;
  Location_col: string;
  Description_col: string;
  ProfileLink_col: string;
  SubmittedBy_col: string;
}

export interface IUserProfile {
  displayName: string;
  jobTitle: string;
  department: string;
  employeeId: string;
  mail: string;
}

export interface IMaintenanceRequestFormData {
  requestedBy?: string;
  requestedBy_key?: string;
  requestedFor: string;
  requestedFor_Title: string;
  requestedFor_key: string;
  
  Category: string;
  SubCategory: string;
  Department: string;
  Location: string;
  Description: string;
  
  files?: any;
}