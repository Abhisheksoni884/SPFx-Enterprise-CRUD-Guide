// Centralized Constants for SPFx Employee Leave Management WebPart

export const LIST_NAMES = {
  EMPLOYEE_LEAVE_REQUESTS: 'Employee Leave Requests',
  LOG_HISTORY: 'LogHistory',
  TEAM_MAPPING: 'Manager_List'
} as const;

export const LEAVE_SELECT_FIELDS = [
  'Id',
  'Title',
  'EmployeeName/Id',
  'EmployeeName/Title',
  'EmployeeName/EMail',
  'LeaveType',
  'StartDate',
  'EndDate',
  'Reason',
  'Status',
  'ManagerComments',
  'AttachmentFiles/FileName',
  'AttachmentFiles/ServerRelativeUrl'
] as const;

export const TEAM_MAPPING_SELECT_FIELDS = [
  'Id',
  'Manager/Id',
  'Manager/Title',
  'Manager/EMail',
  'Member/Id',
  'Member/Title',
  'Member/EMail'
] as const;

export const REASON_MAX_LENGTH = 500;

export const MESSAGES = {
  SUCCESS_SUBMIT: 'Leave request submitted successfully.',
  SUCCESS_UPDATE: 'Leave request updated successfully.',
  SUCCESS_CANCEL: 'Leave request cancelled successfully.',
  SUCCESS_APPROVE: 'Leave request approved successfully.',
  SUCCESS_REJECT: 'Leave request rejected successfully.',
  ERROR_LOAD: 'Unable to load leave requests.',
  ERROR_SAVE: 'Unable to save the leave request.',
  ERROR_CANCEL: 'Unable to cancel the leave request.',
  ERROR_APPROVE: 'Unable to approve leave request.',
  ERROR_REJECT: 'Unable to reject leave request.'
} as const;
