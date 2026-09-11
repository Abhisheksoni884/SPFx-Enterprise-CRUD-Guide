import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/attachments';
import { getSP } from '../../../framework/Services/pnpjsConfig';
import { logError } from '../../../framework/Services/LoggingService';
import { LEAVE_SELECT_FIELDS, LIST_NAMES } from '../../../framework/Constants/Constant';
import { IEmployeeLeaveRequest, ILeaveRequestInput, IPersonaInfo, LeaveStatus, LeaveType } from '../models/IEmployeeLeaveRequest';

// Raw shape returned by SharePoint REST for a list item with EmployeeName and AttachmentFiles expanded.
interface IRawLeaveRequestItem {
  Id: number;
  Title: string;
  EmployeeName?: { Id: number; Title: string; EMail: string };
  LeaveType: LeaveType;
  StartDate: string;
  EndDate: string;
  Reason: string;
  Status: LeaveStatus;
  ManagerComments: string;
  AttachmentFiles?: { results?: Array<{ FileName: string; ServerRelativeUrl: string }> } | Array<{ FileName: string; ServerRelativeUrl: string }>;
}

function mapItem(item: IRawLeaveRequestItem): IEmployeeLeaveRequest {
  let attachments: Array<{ FileName: string; ServerRelativeUrl: string }> = [];
  if (Array.isArray(item.AttachmentFiles)) {
    attachments = item.AttachmentFiles;
  } else if (item.AttachmentFiles && Array.isArray(item.AttachmentFiles.results)) {
    attachments = item.AttachmentFiles.results;
  }

  return {
    Id: item.Id,
    Title: item.Title,
    EmployeeNameId: item.EmployeeName ? item.EmployeeName.Id : 0,
    EmployeeName: item.EmployeeName ? item.EmployeeName.Title : '',
    EmployeeEmail: item.EmployeeName ? item.EmployeeName.EMail : '',
    LeaveType: item.LeaveType,
    StartDate: item.StartDate,
    EndDate: item.EndDate,
    Reason: item.Reason || '',
    Status: item.Status,
    ManagerComments: item.ManagerComments || '',
    AttachmentFiles: attachments
  };
}

/**
 * Service encapsulating SharePoint REST API operations for Employee Leave Requests using pure functions.
 */
export const getAllLeaveRequests = async (
  context: WebPartContext,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<IEmployeeLeaveRequest[]> => {
  try {
    const sp = getSP(context);
    const items: IRawLeaveRequestItem[] = await sp.web.lists.getByTitle(listName).items
      .select(...LEAVE_SELECT_FIELDS)
      .expand('EmployeeName', 'AttachmentFiles')
      .orderBy('Id', false)
      .top(500)();

    return items.map(mapItem);
  } catch (error) {
    await logError(context, 'getAllLeaveRequests', 'LeaveRequestService', error);
    throw error;
  }
};

export const resolveEmployee = async (
  context: WebPartContext,
  loginName: string,
  displayName: string,
  email: string
): Promise<IPersonaInfo> => {
  try {
    const sp = getSP(context);
    const ensured = await sp.web.ensureUser(loginName);
    return { id: ensured.Id, loginName: ensured.LoginName, displayName, email };
  } catch (error) {
    await logError(context, 'resolveEmployee', 'LeaveRequestService', error);
    throw error;
  }
};

export const uploadAttachments = async (
  context: WebPartContext,
  itemId: number,
  files: File[],
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  if (!files || files.length === 0) return;
  try {
    const sp = getSP(context);
    const attachments = sp.web.lists.getByTitle(listName).items.getById(itemId).attachmentFiles;
    for (const file of files) {
      await attachments.add(file.name, file);
    }
  } catch (error) {
    await logError(context, 'uploadAttachments', 'LeaveRequestService', error);
    throw error;
  }
};

export const deleteAttachments = async (
  context: WebPartContext,
  itemId: number,
  fileNames: string[],
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  if (!fileNames || fileNames.length === 0) return;
  try {
    const sp = getSP(context);
    for (const name of fileNames) {
      await sp.web.lists.getByTitle(listName).items.getById(itemId).attachmentFiles.getByName(name).delete();
    }
  } catch (error) {
    await logError(context, 'deleteAttachments', 'LeaveRequestService', error);
    throw error;
  }
};

export const createLeaveRequest = async (
  context: WebPartContext,
  input: ILeaveRequestInput,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<IEmployeeLeaveRequest> => {
  try {
    const sp = getSP(context);
    const result = await sp.web.lists.getByTitle(listName).items.add({
      Title: `${input.employee.displayName} - ${input.leaveType}`,
      EmployeeNameId: input.employee.id,
      LeaveType: input.leaveType,
      StartDate: input.startDate.toISOString(),
      EndDate: input.endDate.toISOString(),
      Reason: input.reason,
      Status: LeaveStatus.Pending
    });

    const resAny = result as { data?: { Id?: number; ID?: number }; Id?: number; ID?: number };
    const newItemId = resAny?.data?.Id || resAny?.data?.ID || resAny?.Id || resAny?.ID || 0;

    if (newItemId > 0 && input.attachments && input.attachments.length > 0) {
      await uploadAttachments(context, newItemId, input.attachments, listName);
    }

    return {
      Id: newItemId,
      Title: `${input.employee.displayName} - ${input.leaveType}`,
      EmployeeNameId: input.employee.id,
      EmployeeName: input.employee.displayName,
      EmployeeEmail: input.employee.email,
      LeaveType: input.leaveType,
      StartDate: input.startDate.toISOString(),
      EndDate: input.endDate.toISOString(),
      Reason: input.reason,
      Status: LeaveStatus.Pending,
      ManagerComments: '',
      AttachmentFiles: input.attachments ? input.attachments.map(f => ({ FileName: f.name, ServerRelativeUrl: '' })) : []
    };
  } catch (error) {
    await logError(context, 'createLeaveRequest', 'LeaveRequestService', error);
    throw error;
  }
};

export const updateLeaveRequest = async (
  context: WebPartContext,
  id: number,
  input: ILeaveRequestInput,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  try {
    const sp = getSP(context);
    await sp.web.lists.getByTitle(listName).items.getById(id).update({
      Title: `${input.employee.displayName} - ${input.leaveType}`,
      EmployeeNameId: input.employee.id,
      LeaveType: input.leaveType,
      StartDate: input.startDate.toISOString(),
      EndDate: input.endDate.toISOString(),
      Reason: input.reason
    });

    if (input.attachmentsToDelete && input.attachmentsToDelete.length > 0) {
      await deleteAttachments(context, id, input.attachmentsToDelete, listName);
    }

    if (input.attachments && input.attachments.length > 0) {
      await uploadAttachments(context, id, input.attachments, listName);
    }
  } catch (error) {
    await logError(context, 'updateLeaveRequest', 'LeaveRequestService', error);
    throw error;
  }
};

export const approveLeaveRequest = async (
  context: WebPartContext,
  id: number,
  managerComments?: string,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  try {
    const sp = getSP(context);
    await sp.web.lists.getByTitle(listName).items.getById(id).update({
      Status: LeaveStatus.Approved,
      ManagerComments: managerComments || ''
    });
  } catch (error) {
    await logError(context, 'approveLeaveRequest', 'LeaveRequestService', error);
    throw error;
  }
};

export const rejectLeaveRequest = async (
  context: WebPartContext,
  id: number,
  managerComments?: string,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  try {
    const sp = getSP(context);
    await sp.web.lists.getByTitle(listName).items.getById(id).update({
      Status: LeaveStatus.Rejected,
      ManagerComments: managerComments || ''
    });
  } catch (error) {
    await logError(context, 'rejectLeaveRequest', 'LeaveRequestService', error);
    throw error;
  }
};

export const cancelLeaveRequest = async (
  context: WebPartContext,
  id: number,
  listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS
): Promise<void> => {
  try {
    const sp = getSP(context);
    const item = sp.web.lists.getByTitle(listName).items.getById(id);
    if (typeof item.recycle === 'function') {
      await item.recycle();
    } else {
      await item.delete();
    }
  } catch (error) {
    await logError(context, 'cancelLeaveRequest', 'LeaveRequestService', error);
    throw error;
  }
};

export default {
  getAllLeaveRequests,
  resolveEmployee,
  uploadAttachments,
  deleteAttachments,
  createLeaveRequest,
  updateLeaveRequest,
  approveLeaveRequest,
  rejectLeaveRequest,
  cancelLeaveRequest
};

