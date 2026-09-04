import * as React from 'react';
import { useMemo, useState } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  SearchBox,
  Dropdown,
  IDropdownOption,
  IconButton,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  Persona,
  PersonaSize,
  Icon,
  Label,
  TooltipHost
} from '@fluentui/react';
import { IEmployeeLeaveRequest, LeaveStatus, STATUS_FILTER_OPTIONS } from '../models/IEmployeeLeaveRequest';
import { formatDate } from '../../../framework/Utilities/Utilities';
import styles from './CrudOperation.module.scss';

export interface ILeaveRequestListProps {
  items: IEmployeeLeaveRequest[];
  isLoading: boolean;
  onEdit: (item: IEmployeeLeaveRequest) => void;
  onDelete: (item: IEmployeeLeaveRequest) => void;
  onApprove?: (item: IEmployeeLeaveRequest) => void;
  onReject?: (item: IEmployeeLeaveRequest) => void;
  isManagerView?: boolean;
}

const STATUS_OPTIONS: IDropdownOption[] = STATUS_FILTER_OPTIONS.map((status) => ({ key: status, text: status }));

function getStatusClassName(status: LeaveStatus): string {
  switch (status) {
    case LeaveStatus.Approved:
      return styles.statusApproved;
    case LeaveStatus.Rejected:
      return styles.statusRejected;
    case LeaveStatus.Pending:
    default:
      return styles.statusPending;
  }
}

function getStatusIconName(status: LeaveStatus): string {
  switch (status) {
    case LeaveStatus.Approved:
      return 'CheckMark';
    case LeaveStatus.Rejected:
      return 'Cancel';
    case LeaveStatus.Pending:
    default:
      return 'Clock';
  }
}

// Table view for the "Employee Leave Requests" list: search, status filter and row actions.
const LeaveRequestList: React.FC<ILeaveRequestListProps> = (props) => {
  const { items, isLoading, onEdit, onDelete, onApprove, onReject, isManagerView } = props;
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('All');

  const filteredItems = useMemo(() => {
    const search = searchText.trim().toLowerCase();
    return items.filter((item) => {
      const matchesStatus = statusFilter === 'All' || item.Status === statusFilter;
      const matchesSearch =
        search.length === 0 ||
        item.EmployeeName.toLowerCase().indexOf(search) !== -1 ||
        item.LeaveType.toLowerCase().indexOf(search) !== -1;
      return matchesStatus && matchesSearch;
    });
  }, [items, searchText, statusFilter]);

  const columns: IColumn[] = [
    {
      key: 'employee',
      name: 'Employee',
      minWidth: 120,
      maxWidth: 200,
      onRender: (item: IEmployeeLeaveRequest) => (
        <Persona text={item.EmployeeName} secondaryText={item.EmployeeEmail} size={PersonaSize.size24} className={styles.employeeName} />
      )
    },
    {
      key: 'leaveType',
      name: 'Leave Type',
      minWidth: 110,
      maxWidth: 130,
      onRender: (item: IEmployeeLeaveRequest) => (
        <span className={styles.leaveTypeBadge}>
          <Icon iconName="Calendar" />
          {item.LeaveType}
        </span>
      )
    },
    {
      key: 'startDate',
      name: 'Start Date',
      minWidth: 90,
      maxWidth: 110,
      onRender: (item: IEmployeeLeaveRequest) => <span>{formatDate(item.StartDate)}</span>
    },
    {
      key: 'endDate',
      name: 'End Date',
      minWidth: 90,
      maxWidth: 110,
      onRender: (item: IEmployeeLeaveRequest) => <span>{formatDate(item.EndDate)}</span>
    },
    { key: 'reason', name: 'Reason', fieldName: 'Reason', minWidth: 100, maxWidth: 200, isMultiline: true },
    {
      key: 'comments',
      name: 'Manager Comments',
      fieldName: 'ManagerComments',
      minWidth: 110,
      maxWidth: 200,
      isMultiline: true,
      onRender: (item: IEmployeeLeaveRequest) => (
        item.ManagerComments ? (
          <TooltipHost content={item.ManagerComments}>
            <span style={{ fontStyle: 'italic', color: '#605e5c' }}>{item.ManagerComments}</span>
          </TooltipHost>
        ) : (
          <span style={{ color: '#a19f9d' }}>-</span>
        )
      )
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 90,
      maxWidth: 100,
      onRender: (item: IEmployeeLeaveRequest) => (
        <span className={`${styles.statusBadge} ${getStatusClassName(item.Status)}`}>
          <Icon iconName={getStatusIconName(item.Status)} />
          {item.Status}
        </span>
      )
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 130,
      maxWidth: 150,
      onRender: (item: IEmployeeLeaveRequest) => {
        const isPending = item.Status === LeaveStatus.Pending;

        if (isManagerView) {
          return (
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              <IconButton
                iconProps={{ iconName: 'Accept' }}
                title={isPending ? 'Approve Request' : 'Only pending requests can be approved'}
                ariaLabel="Approve"
                disabled={!isPending}
                style={{ color: isPending ? '#107c10' : undefined }}
                onClick={() => onApprove && onApprove(item)}
              />
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title={isPending ? 'Reject Request' : 'Only pending requests can be rejected'}
                ariaLabel="Reject"
                disabled={!isPending}
                style={{ color: isPending ? '#a4262c' : undefined }}
                onClick={() => onReject && onReject(item)}
              />
            </Stack>
          );
        }

        return (
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title={isPending ? 'Edit' : 'Only pending requests can be edited'}
              ariaLabel="Edit"
              disabled={!isPending}
              onClick={() => onEdit(item)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title={isPending ? 'Delete' : 'Only pending requests can be cancelled'}
              ariaLabel="Delete"
              disabled={!isPending}
              onClick={() => onDelete(item)}
            />
          </Stack>
        );
      }
    }
  ];

  return (
    <div>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }} className={styles.toolbarBar}>
        <Stack>
          <Label>Search</Label>
          <SearchBox
            placeholder="By employee or leave type"
            value={searchText}
            onChange={(_, value) => setSearchText(value || '')}
            styles={{ root: { width: 280 } }}
          />
        </Stack>
        <Dropdown
          label="Status"
          selectedKey={statusFilter}
          options={STATUS_OPTIONS}
          onChange={(_, option) => setStatusFilter(option ? (option.key as string) : 'All')}
          styles={{ root: { width: 160 } }}
        />
      </Stack>

      {isLoading && <Spinner size={SpinnerSize.large} label="Loading leave requests..." />}

      {!isLoading && filteredItems.length === 0 && (
        <div className={styles.emptyState}>
          <Icon iconName="EventDate" className={styles.emptyStateIcon} />
          <Text variant="mediumPlus">No leave requests found.</Text>
        </div>
      )}

      {!isLoading && filteredItems.length > 0 && (
        <div className={styles.tableWrapper}>
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            getKey={(item: IEmployeeLeaveRequest) => item.Id.toString()}
          />
        </div>
      )}
    </div>
  );
};

export default LeaveRequestList;
