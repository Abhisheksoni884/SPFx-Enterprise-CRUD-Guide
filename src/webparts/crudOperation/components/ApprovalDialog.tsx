import * as React from 'react';
import { useState } from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { IEmployeeLeaveRequest, LeaveStatus } from '../models/IEmployeeLeaveRequest';

export interface IApprovalDialogProps {
  item?: IEmployeeLeaveRequest;
  action: 'Approve' | 'Reject';
  isOpen: boolean;
  isSubmitting: boolean;
  onDismiss: () => void;
  onConfirm: (item: IEmployeeLeaveRequest, comments: string) => void;
}

const ApprovalDialog: React.FC<IApprovalDialogProps> = (props) => {
  const { item, action, isOpen, isSubmitting, onDismiss, onConfirm } = props;
  const [comments, setComments] = useState<string>('');

  if (!item) {
    return null;
  }

  const isApprove = action === 'Approve';

  const handleConfirm = (): void => {
    onConfirm(item, comments);
    setComments('');
  };

  return (
    <Dialog
      hidden={!isOpen}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: `${action} Leave Request`,
        subText: `Are you sure you want to ${action.toLowerCase()} the leave request for ${item.EmployeeName}?`
      }}
    >
      <MessageBar messageBarType={isApprove ? MessageBarType.info : MessageBarType.warning} style={{ marginBottom: 12 }}>
        Status will be set to <strong>{isApprove ? LeaveStatus.Approved : LeaveStatus.Rejected}</strong>.
      </MessageBar>

      <TextField
        label="Manager Comments (Optional)"
        multiline
        rows={3}
        value={comments}
        onChange={(_, val) => setComments(val || '')}
        placeholder="Add remarks or justification..."
      />

      <DialogFooter>
        <PrimaryButton
          text={action}
          onClick={handleConfirm}
          disabled={isSubmitting}
          style={{ backgroundColor: isApprove ? '#107c10' : '#a4262c', borderColor: isApprove ? '#107c10' : '#a4262c' }}
        />
        <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSubmitting} />
      </DialogFooter>
    </Dialog>
  );
};

export default ApprovalDialog;
