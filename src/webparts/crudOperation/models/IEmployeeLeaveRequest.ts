import { REASON_MAX_LENGTH as CONSTANT_REASON_MAX_LENGTH } from '../../../framework/Constants/Constant';

export enum LeaveType {
  Casual = 'Casual Leave',
  Sick = 'Sick Leave',
  Annual = 'Annual Leave',
  Other = 'Other'
}

export enum LeaveStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected'
}

// Resolved person, used by the People Picker and the SharePoint "Employee Name" field.
export interface IPersonaInfo {
  id: number;
  loginName: string;
  displayName: string;
  email: string;
}

export interface IAttachmentInfo {
  FileName: string;
  ServerRelativeUrl: string;
}

// Shape returned when reading items from the "Employee Leave Requests" list.
export interface IEmployeeLeaveRequest {
  Id: number;
  Title: string;
  EmployeeNameId: number;
  EmployeeName: string;
  EmployeeEmail: string;
  LeaveType: LeaveType;
  StartDate: string;
  EndDate: string;
  Reason: string;
  Status: LeaveStatus;
  ManagerComments: string;
  AttachmentFiles?: IAttachmentInfo[];
}

// Shape used by the create/edit form before it is sent to SharePoint.
export interface ILeaveRequestInput {
  employee: IPersonaInfo;
  leaveType: LeaveType;
  startDate: Date;
  endDate: Date;
  reason: string;
  attachments?: File[];
  attachmentsToDelete?: string[];
}

export const REASON_MAX_LENGTH = CONSTANT_REASON_MAX_LENGTH;

export const STATUS_FILTER_OPTIONS: Array<LeaveStatus | 'All'> = ['All', LeaveStatus.Pending, LeaveStatus.Approved, LeaveStatus.Rejected];
