import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  DatePicker,
  TextField,
  Stack
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './CrudOperation.module.scss';
import PeoplePickerField from './PeoplePickerField';
import {
  IEmployeeLeaveRequest,
  ILeaveRequestInput,
  IPersonaInfo,
  LeaveType,
  REASON_MAX_LENGTH
} from '../models/IEmployeeLeaveRequest';
import { isDateBefore } from '../../../framework/Utilities/Utilities';

export interface ILeaveRequestFormProps {
  context: WebPartContext;
  isOpen: boolean;
  isSaving: boolean;
  // Present when editing an existing request; absent when creating a new one.
  editingItem?: IEmployeeLeaveRequest;
  onDismiss: () => void;
  onSubmit: (input: ILeaveRequestInput, id?: number) => void;
}

interface IFormErrors {
  employee?: string;
  leaveType?: string;
  startDate?: string;
  endDate?: string;
  reason?: string;
}

const LEAVE_TYPE_OPTIONS: IDropdownOption[] = Object.keys(LeaveType).map((key) => {
  const value = LeaveType[key as keyof typeof LeaveType];
  return { key: value, text: value };
});

function toDate(value?: string): Date | undefined {
  return value ? new Date(value) : undefined;
}

const LeaveRequestForm: React.FC<ILeaveRequestFormProps> = (props) => {
  const { context, isOpen, isSaving, editingItem, onDismiss, onSubmit } = props;
  const isEditMode = !!editingItem;

  const [employee, setEmployee] = useState<IPersonaInfo | undefined>(undefined);
  const [leaveType, setLeaveType] = useState<LeaveType | undefined>(undefined);
  const [startDate, setStartDate] = useState<Date | undefined>(undefined);
  const [endDate, setEndDate] = useState<Date | undefined>(undefined);
  const [reason, setReason] = useState<string>('');
  const [errors, setErrors] = useState<IFormErrors>({});

  // Re-seed the form fields every time the panel is (re)opened.
  useEffect(() => {
    if (!isOpen) {
      return;
    }

    setErrors({});

    if (editingItem) {
      setEmployee({
        id: editingItem.EmployeeNameId,
        loginName: '',
        displayName: editingItem.EmployeeName,
        email: editingItem.EmployeeEmail
      });
      setLeaveType(editingItem.LeaveType);
      setStartDate(toDate(editingItem.StartDate));
      setEndDate(toDate(editingItem.EndDate));
      setReason(editingItem.Reason || '');
    } else {
      const currentUser = context.pageContext.user;
      setEmployee({
        id: 0,
        loginName: currentUser.loginName,
        displayName: currentUser.displayName,
        email: currentUser.email
      });
      setLeaveType(undefined);
      setStartDate(undefined);
      setEndDate(undefined);
      setReason('');
    }
  }, [isOpen, editingItem, context]);

  const validate = (): IFormErrors => {
    const nextErrors: IFormErrors = {};

    if (!employee) {
      nextErrors.employee = 'Employee Name is required.';
    }
    if (!leaveType) {
      nextErrors.leaveType = 'Leave Type is required.';
    }
    if (!startDate) {
      nextErrors.startDate = 'Start Date is required.';
    }
    if (!endDate) {
      nextErrors.endDate = 'End Date is required.';
    }
    if (startDate && endDate && isDateBefore(startDate, endDate)) {
      nextErrors.endDate = 'End Date cannot be earlier than Start Date.';
    }
    if (reason && reason.length > REASON_MAX_LENGTH) {
      nextErrors.reason = `Reason cannot exceed ${REASON_MAX_LENGTH} characters.`;
    }

    return nextErrors;
  };

  const handleSubmit = (): void => {
    const nextErrors = validate();
    setErrors(nextErrors);

    if (Object.keys(nextErrors).length > 0 || !employee || !leaveType || !startDate || !endDate) {
      return;
    }

    onSubmit({ employee, leaveType, startDate, endDate, reason }, editingItem ? editingItem.Id : undefined);
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={isEditMode ? 'Edit Leave Request' : 'Apply for Leave'}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text={isEditMode ? 'Update' : 'Submit'} onClick={handleSubmit} disabled={isSaving} />
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSaving} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <div className={styles.formSection}>
          <p className={styles.formSectionTitle}>Requester</p>
          <PeoplePickerField
            context={context}
            label="Employee Name"
            required
            errorMessage={errors.employee}
            selectedPerson={employee}
            onChange={setEmployee}
          />
        </div>

        <div className={styles.formSection}>
          <p className={styles.formSectionTitle}>Leave Details</p>
          <Dropdown
            label="Leave Type"
            required
            selectedKey={leaveType}
            options={LEAVE_TYPE_OPTIONS}
            errorMessage={errors.leaveType}
            onChange={(_, option) => setLeaveType(option ? (option.key as LeaveType) : undefined)}
          />
        </div>

        <div className={styles.dateRow}>
          <div>
            <DatePicker
              label="Start Date"
              isRequired
              value={startDate}
              onSelectDate={(date) => setStartDate(date || undefined)}
              formatDate={(date) => (date ? date.toLocaleDateString() : '')}
            />
            {errors.startDate && <span className={styles.fieldError}>{errors.startDate}</span>}
          </div>

          <div>
            <DatePicker
              label="End Date"
              isRequired
              value={endDate}
              minDate={startDate}
              onSelectDate={(date) => setEndDate(date || undefined)}
              formatDate={(date) => (date ? date.toLocaleDateString() : '')}
            />
            {errors.endDate && <span className={styles.fieldError}>{errors.endDate}</span>}
          </div>
        </div>

        <TextField
          label="Reason"
          multiline
          rows={4}
          resizable={false}
          value={reason}
          maxLength={REASON_MAX_LENGTH}
          description={`${reason.length}/${REASON_MAX_LENGTH} characters`}
          errorMessage={errors.reason}
          onChange={(_, value) => setReason(value || '')}
        />
      </Stack>
    </Panel>
  );
};

export default LeaveRequestForm;
