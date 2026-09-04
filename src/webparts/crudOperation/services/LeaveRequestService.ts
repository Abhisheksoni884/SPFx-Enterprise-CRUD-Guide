import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IItems } from '@pnp/sp/items';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../framework/Services/pnpjsConfig';
import LoggingService from '../../../framework/Services/LoggingService';
import { LEAVE_SELECT_FIELDS, LIST_NAMES } from '../../../framework/Constants/Constant';
import { IEmployeeLeaveRequest, ILeaveRequestInput, IPersonaInfo, LeaveStatus, LeaveType } from '../models/IEmployeeLeaveRequest';

// Raw shape returned by SharePoint REST for a list item with the EmployeeName person field expanded.
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
}

function mapItem(item: IRawLeaveRequestItem): IEmployeeLeaveRequest {
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
    ManagerComments: item.ManagerComments || ''
  };
}

/**
 * Service encapsulating SharePoint REST API operations for Employee Leave Requests.
 * Utilizes the PnP JS singleton and logs operational failures to SharePoint LogHistory list.
 */
export default class LeaveRequestService {
  private readonly _sp: SPFI;

  constructor(private readonly context: WebPartContext, private readonly listName: string = LIST_NAMES.EMPLOYEE_LEAVE_REQUESTS) {
    this._sp = getSP(context);
  }

  private get _items(): IItems {
    return this._sp.web.lists.getByTitle(this.listName).items;
  }

  public async getAll(): Promise<IEmployeeLeaveRequest[]> {
    try {
      const items: IRawLeaveRequestItem[] = await this._items
        .select(...LEAVE_SELECT_FIELDS)
        .expand('EmployeeName')
        .orderBy('Id', false)
        .top(500)();

      return items.map(mapItem);
    } catch (error) {
      await LoggingService.logError(this.context, 'getAll', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async resolveEmployee(loginName: string, displayName: string, email: string): Promise<IPersonaInfo> {
    try {
      const ensured = await this._sp.web.ensureUser(loginName);
      return { id: ensured.Id, loginName: ensured.LoginName, displayName, email };
    } catch (error) {
      await LoggingService.logError(this.context, 'resolveEmployee', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async create(input: ILeaveRequestInput): Promise<IEmployeeLeaveRequest> {
    try {
      const result = await this._items.add({
        Title: `${input.employee.displayName} - ${input.leaveType}`,
        EmployeeNameId: input.employee.id,
        LeaveType: input.leaveType,
        StartDate: input.startDate.toISOString(),
        EndDate: input.endDate.toISOString(),
        Reason: input.reason,
        Status: LeaveStatus.Pending
      });

      // Safely extract Id from PnP JS v4 result object across different response structures
      const resAny = result as { data?: { Id?: number; ID?: number }; Id?: number; ID?: number };
      const newItemId = resAny?.data?.Id || resAny?.data?.ID || resAny?.Id || resAny?.ID || 0;

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
        ManagerComments: ''
      };
    } catch (error) {
      await LoggingService.logError(this.context, 'create', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async update(id: number, input: ILeaveRequestInput): Promise<void> {
    try {
      await this._items.getById(id).update({
        Title: `${input.employee.displayName} - ${input.leaveType}`,
        EmployeeNameId: input.employee.id,
        LeaveType: input.leaveType,
        StartDate: input.startDate.toISOString(),
        EndDate: input.endDate.toISOString(),
        Reason: input.reason
      });
    } catch (error) {
      await LoggingService.logError(this.context, 'update', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async approveRequest(id: number, managerComments?: string): Promise<void> {
    try {
      await this._items.getById(id).update({
        Status: LeaveStatus.Approved,
        ManagerComments: managerComments || ''
      });
    } catch (error) {
      await LoggingService.logError(this.context, 'approveRequest', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async rejectRequest(id: number, managerComments?: string): Promise<void> {
    try {
      await this._items.getById(id).update({
        Status: LeaveStatus.Rejected,
        ManagerComments: managerComments || ''
      });
    } catch (error) {
      await LoggingService.logError(this.context, 'rejectRequest', 'LeaveRequestService', error);
      throw error;
    }
  }

  public async remove(id: number): Promise<void> {
    try {
      const item = this._items.getById(id);
      if (typeof item.recycle === 'function') {
        await item.recycle();
      } else {
        await item.delete();
      }
    } catch (error) {
      await LoggingService.logError(this.context, 'remove', 'LeaveRequestService', error);
      throw error;
    }
  }
}
