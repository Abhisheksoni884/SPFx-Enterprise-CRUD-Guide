import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
  Icon,
  Stack,
  Pivot,
  PivotItem
} from '@fluentui/react';
import styles from './CrudOperation.module.scss';
import type { ICrudOperationProps } from './ICrudOperationProps';
import LeaveRequestList from './LeaveRequestList';
import LeaveRequestForm from './LeaveRequestForm';
import ApprovalDialog from './ApprovalDialog';
import LeaveRequestService from '../services/LeaveRequestService';
import TeamMappingService from '../../../framework/Services/TeamMappingService';
import GraphEmailService from '../../../framework/Services/GraphEmailService';
import { IEmployeeLeaveRequest, ILeaveRequestInput, LeaveStatus } from '../models/IEmployeeLeaveRequest';
import { MESSAGES } from '../../../framework/Constants/Constant';

interface IManagedTeam {
  emails: string[];
  titles: string[];
  ids: number[];
}

const CrudOperation: React.FC<ICrudOperationProps> = (props) => {
  const { context, listName } = props;
  const service = useMemo(() => new LeaveRequestService(context, listName), [context, listName]);
  const teamMappingService = useMemo(() => new TeamMappingService(context), [context]);

  const userEmail = context.pageContext.user.email || '';
  const userLogin = context.pageContext.user.loginName || '';
  const userName = context.pageContext.user.displayName || '';

  const [items, setItems] = useState<IEmployeeLeaveRequest[]>([]);
  const [managedTeam, setManagedTeam] = useState<IManagedTeam>({ emails: [], titles: [], ids: [] });
  const [isManager, setIsManager] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<string>('myRequests');

  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [editingItem, setEditingItem] = useState<IEmployeeLeaveRequest | undefined>(undefined);
  const [pendingDelete, setPendingDelete] = useState<IEmployeeLeaveRequest | undefined>(undefined);
  
  // Manager Approval State
  const [approvalItem, setApprovalItem] = useState<IEmployeeLeaveRequest | undefined>(undefined);
  const [approvalAction, setApprovalAction] = useState<'Approve' | 'Reject'>('Approve');
  const [isApprovalSubmitting, setIsApprovalSubmitting] = useState<boolean>(false);

  const [successMessage, setSuccessMessage] = useState<string | undefined>(undefined);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);

  // Check Manager Role & Managed Team Members on Mount
  useEffect(() => {
    const initRoleAndTeam = async (): Promise<void> => {
      try {
        const userIsManager = await teamMappingService.isUserAManager(userEmail, userLogin, userName);
        setIsManager(userIsManager);

        if (userIsManager) {
          const team = await teamMappingService.getManagedTeamMembers(userEmail, userLogin, userName);
          setManagedTeam(team);
        }
      } catch (err) {
        console.warn('Failed to load team mappings:', err);
      }
    };

    initRoleAndTeam().catch(() => undefined);
  }, [teamMappingService, userEmail, userLogin, userName]);

  const loadItems = async (): Promise<void> => {
    setIsLoading(true);
    setErrorMessage(undefined);
    try {
      const results = await service.getAll();
      setItems(results);
    } catch (error) {
      setErrorMessage(`${MESSAGES.ERROR_LOAD}: ${(error as Error).message}`);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadItems().catch(() => {
      /* handled inside loadItems */
    });
  }, [service]);

  // Filter items based on active Tab (My Requests vs Manager Dashboard)
  const displayedItems = useMemo(() => {
    if (activeTab === 'managerDashboard') {
      const emailSet = new Set(managedTeam.emails);
      const titleSet = new Set(managedTeam.titles);
      const idSet = new Set(managedTeam.ids);

      return items.filter((item) => {
        const matchesEmail = item.EmployeeEmail && emailSet.has(item.EmployeeEmail.toLowerCase().trim());
        const matchesTitle = item.EmployeeName && titleSet.has(item.EmployeeName.toLowerCase().trim());
        const matchesId = item.EmployeeNameId && idSet.has(item.EmployeeNameId);
        return matchesEmail || matchesTitle || matchesId;
      });
    }

    // "My Leave Requests" tab
    const emailLower = userEmail.toLowerCase().trim();
    const loginLower = userLogin.toLowerCase().trim();
    const nameLower = userName.toLowerCase().trim();

    return items.filter((item) => {
      const empEmail = (item.EmployeeEmail || '').toLowerCase().trim();
      const empName = (item.EmployeeName || '').toLowerCase().trim();
      return (
        (emailLower && empEmail === emailLower) ||
        (nameLower && empName === nameLower) ||
        (loginLower && empEmail.indexOf(loginLower) !== -1)
      );
    });
  }, [items, activeTab, managedTeam, userEmail, userLogin, userName]);

  const summary = useMemo(
    () => ({
      pending: displayedItems.filter((item) => item.Status === LeaveStatus.Pending).length,
      approved: displayedItems.filter((item) => item.Status === LeaveStatus.Approved).length,
      rejected: displayedItems.filter((item) => item.Status === LeaveStatus.Rejected).length
    }),
    [displayedItems]
  );

  const openNewForm = (): void => {
    setEditingItem(undefined);
    setIsPanelOpen(true);
  };

  const openEditForm = (item: IEmployeeLeaveRequest): void => {
    setEditingItem(item);
    setIsPanelOpen(true);
  };

  const closeForm = (): void => {
    setIsPanelOpen(false);
  };

  const handleSubmit = async (input: ILeaveRequestInput, id?: number): Promise<void> => {
    setIsSaving(true);
    setErrorMessage(undefined);
    try {
      // id === 0 means freshly picked persona not yet resolved to a SharePoint user ID.
      const resolvedEmployee =
        input.employee.id === 0
          ? await service.resolveEmployee(input.employee.loginName, input.employee.displayName, input.employee.email)
          : input.employee;
      const resolvedInput: ILeaveRequestInput = { ...input, employee: resolvedEmployee };

      if (id) {
        await service.update(id, resolvedInput);
        setSuccessMessage(MESSAGES.SUCCESS_UPDATE);
      } else {
        const createdItem = await service.create(resolvedInput);
        setSuccessMessage(MESSAGES.SUCCESS_SUBMIT);

        // Send Email Notification to Manager via MS Graph API
        try {
          const managerInfo = await teamMappingService.getManagerForEmployee(
            resolvedEmployee.email,
            resolvedEmployee.loginName,
            resolvedEmployee.displayName
          );
          if (managerInfo && managerInfo.email) {
            await GraphEmailService.sendLeaveNotificationToManager(
              context,
              managerInfo.email,
              managerInfo.displayName,
              createdItem
            );
          }
        } catch (emailErr) {
          console.warn('Failed to send manager email notification:', emailErr);
        }
      }

      setIsPanelOpen(false);
      await loadItems();
    } catch (error) {
      setErrorMessage(`${MESSAGES.ERROR_SAVE}: ${(error as Error).message}`);
    } finally {
      setIsSaving(false);
    }
  };

  // Delete Handlers
  const requestDelete = (item: IEmployeeLeaveRequest): void => {
    setPendingDelete(item);
  };

  const confirmDelete = async (): Promise<void> => {
    if (!pendingDelete) {
      return;
    }
    setErrorMessage(undefined);
    try {
      await service.remove(pendingDelete.Id);
      setSuccessMessage(MESSAGES.SUCCESS_CANCEL);
      setPendingDelete(undefined);
      await loadItems();
    } catch (error) {
      setErrorMessage(`${MESSAGES.ERROR_CANCEL}: ${(error as Error).message}`);
      setPendingDelete(undefined);
    }
  };

  // Manager Approval/Rejection Handlers
  const handleOpenApprove = (item: IEmployeeLeaveRequest): void => {
    setApprovalItem(item);
    setApprovalAction('Approve');
  };

  const handleOpenReject = (item: IEmployeeLeaveRequest): void => {
    setApprovalItem(item);
    setApprovalAction('Reject');
  };

  const handleConfirmApproval = async (item: IEmployeeLeaveRequest, comments: string): Promise<void> => {
    setIsApprovalSubmitting(true);
    setErrorMessage(undefined);
    try {
      if (approvalAction === 'Approve') {
        await service.approveRequest(item.Id, comments);
        setSuccessMessage(MESSAGES.SUCCESS_APPROVE);
      } else {
        await service.rejectRequest(item.Id, comments);
        setSuccessMessage(MESSAGES.SUCCESS_REJECT);
      }

      setApprovalItem(undefined);
      await loadItems();
    } catch (error) {
      setErrorMessage(
        approvalAction === 'Approve'
          ? `${MESSAGES.ERROR_APPROVE}: ${(error as Error).message}`
          : `${MESSAGES.ERROR_REJECT}: ${(error as Error).message}`
      );
    } finally {
      setIsApprovalSubmitting(false);
    }
  };

  return (
    <section className={styles.crudOperation}>
      <div className={styles.card}>
        <div className={styles.headerRow}>
          <div className={styles.headerLeft}>
            <span className={styles.headerIcon}>
              <Icon iconName="Calendar" />
            </span>
            <div>
              <h2 className={styles.headerTitle}>Employee Leave Requests</h2>
              <p className={styles.headerSubtitle}>Submit, track and manage leave applications</p>
            </div>
          </div>
          <PrimaryButton text="Apply for Leave" iconProps={{ iconName: 'Add' }} onClick={openNewForm} />
        </div>

        <Pivot
          selectedKey={activeTab}
          onLinkClick={(item) => setActiveTab(item?.props.itemKey || 'myRequests')}
          style={{ marginBottom: 16 }}
        >
          <PivotItem headerText="My Leave Requests" itemKey="myRequests" itemIcon="UserEvent" />
          {isManager && (
            <PivotItem headerText="Manager Dashboard (Team Requests)" itemKey="managerDashboard" itemIcon="People" />
          )}
        </Pivot>

        <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.summaryStack}>
          <span className={`${styles.summaryPill} ${styles.summaryPillPending}`}>{summary.pending} Pending</span>
          <span className={`${styles.summaryPill} ${styles.summaryPillApproved}`}>{summary.approved} Approved</span>
          <span className={`${styles.summaryPill} ${styles.summaryPillRejected}`}>{summary.rejected} Rejected</span>
        </Stack>

        {successMessage && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMessage(undefined)} isMultiline={false}>
            {successMessage}
          </MessageBar>
        )}
        {errorMessage && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)} isMultiline={false}>
            {errorMessage}
          </MessageBar>
        )}

        <LeaveRequestList
          items={displayedItems}
          isLoading={isLoading}
          onEdit={openEditForm}
          onDelete={requestDelete}
          onApprove={handleOpenApprove}
          onReject={handleOpenReject}
          isManagerView={activeTab === 'managerDashboard'}
        />
      </div>

      <LeaveRequestForm
        context={context}
        isOpen={isPanelOpen}
        isSaving={isSaving}
        editingItem={editingItem}
        onDismiss={closeForm}
        onSubmit={handleSubmit}
      />

      {/* Cancel Confirmation Dialog */}
      <Dialog
        hidden={!pendingDelete}
        onDismiss={() => setPendingDelete(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Cancel Leave Request',
          subText: 'Are you sure you want to cancel this leave request?'
        }}
      >
        <DialogFooter>
          <PrimaryButton text="Yes, cancel it" onClick={() => { confirmDelete().catch(() => undefined); }} />
          <DefaultButton text="No" onClick={() => setPendingDelete(undefined)} />
        </DialogFooter>
      </Dialog>

      {/* Manager Approval / Rejection Dialog */}
      <ApprovalDialog
        isOpen={!!approvalItem}
        item={approvalItem}
        action={approvalAction}
        isSubmitting={isApprovalSubmitting}
        onDismiss={() => setApprovalItem(undefined)}
        onConfirm={handleConfirmApproval}
      />
    </section>
  );
};

export default CrudOperation;
