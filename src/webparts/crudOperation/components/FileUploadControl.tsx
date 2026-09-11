import * as React from 'react';
import { useRef, useState } from 'react';
import {
  Icon,
  Label,
  Text,
  IconButton,
  Stack,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { IAttachmentInfo } from '../models/IEmployeeLeaveRequest';

export interface IFileUploadControlProps {
  newFiles: File[];
  existingFiles?: IAttachmentInfo[];
  onNewFilesChange: (files: File[]) => void;
  onRemoveExistingFile?: (fileName: string) => void;
  maxFiles?: number;
  maxSizeBytes?: number;
  disabled?: boolean;
}

const DEFAULT_MAX_FILES = 5;
const DEFAULT_MAX_SIZE_BYTES = 10 * 1024 * 1024; // 10 MB

function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}

export const FileUploadControl: React.FC<IFileUploadControlProps> = (props) => {
  const {
    newFiles,
    existingFiles = [],
    onNewFilesChange,
    onRemoveExistingFile,
    maxFiles = DEFAULT_MAX_FILES,
    maxSizeBytes = DEFAULT_MAX_SIZE_BYTES,
    disabled = false
  } = props;

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragOver, setIsDragOver] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);

  const currentTotalCount = newFiles.length + existingFiles.length;

  const validateAndAddFiles = (incomingFiles: FileList | File[]): void => {
    setErrorMessage(undefined);
    const addedFiles: File[] = [];
    let err: string | undefined = undefined;

    Array.from(incomingFiles).forEach((file) => {
      if (currentTotalCount + addedFiles.length >= maxFiles) {
        err = `Maximum limit of ${maxFiles} attachments reached.`;
        return;
      }
      if (file.size > maxSizeBytes) {
        err = `File "${file.name}" exceeds the maximum allowed size of ${formatFileSize(maxSizeBytes)}.`;
        return;
      }
      // Check for duplicates in newFiles or existingFiles
      const isDuplicateNew = newFiles.some((f) => f.name.toLowerCase() === file.name.toLowerCase());
      const isDuplicateExisting = existingFiles.some((f) => f.FileName.toLowerCase() === file.name.toLowerCase());
      if (isDuplicateNew || isDuplicateExisting) {
        err = `File "${file.name}" is already attached.`;
        return;
      }
      addedFiles.push(file);
    });

    if (err) {
      setErrorMessage(err);
    }

    if (addedFiles.length > 0) {
      onNewFilesChange([...newFiles, ...addedFiles]);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files && e.target.files.length > 0) {
      validateAndAddFiles(e.target.files);
      e.target.value = ''; // Reset input selection
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) {
      setIsDragOver(true);
    }
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    if (!disabled && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      validateAndAddFiles(e.dataTransfer.files);
    }
  };

  const handleRemoveNewFile = (index: number): void => {
    const updated = newFiles.filter((_, i) => i !== index);
    onNewFilesChange(updated);
    setErrorMessage(undefined);
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Label>Attachments (Optional)</Label>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)}>
          {errorMessage}
        </MessageBar>
      )}

      {/* Drag & Drop Zone */}
      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => !disabled && fileInputRef.current?.click()}
        style={{
          border: `2px dashed ${isDragOver ? '#0078d4' : '#ccc'}`,
          borderRadius: 4,
          padding: '16px 20px',
          textAlign: 'center',
          backgroundColor: isDragOver ? '#f0f8ff' : '#fafafa',
          cursor: disabled ? 'not-allowed' : 'pointer',
          transition: 'all 0.2s ease-in-out'
        }}
      >
        <input
          ref={fileInputRef}
          type="file"
          multiple
          style={{ display: 'none' }}
          onChange={handleFileInputChange}
          disabled={disabled}
        />
        <Stack horizontal verticalAlign="center" horizontalAlign="center" tokens={{ childrenGap: 8 }}>
          <Icon iconName="Attach" style={{ fontSize: 20, color: '#0078d4' }} />
          <Text variant="medium">
            <strong>Click to upload</strong> or drag and drop files here
          </Text>
        </Stack>
        <Text variant="small" style={{ color: '#666', display: 'block', marginTop: 4 }}>
          Max {maxFiles} files (up to {formatFileSize(maxSizeBytes)} each)
        </Text>
      </div>

      {/* File List Preview */}
      {(existingFiles.length > 0 || newFiles.length > 0) && (
        <Stack tokens={{ childrenGap: 6 }} style={{ marginTop: 8 }}>
          {/* Existing Attachments */}
          {existingFiles.map((file) => (
            <Stack
              key={file.FileName}
              horizontal
              verticalAlign="center"
              horizontalAlign="space-between"
              style={{
                padding: '6px 12px',
                backgroundColor: '#f3f2f1',
                borderRadius: 4,
                borderLeft: '3px solid #0078d4'
              }}
            >
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Document" style={{ color: '#0078d4' }} />
                <Text variant="medium">{file.FileName}</Text>
                {file.ServerRelativeUrl && (
                  <a
                    href={file.ServerRelativeUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{ fontSize: 12, color: '#0078d4', marginLeft: 8 }}
                    onClick={(e) => e.stopPropagation()}
                  >
                    View
                  </a>
                )}
              </Stack>
              {onRemoveExistingFile && !disabled && (
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  title="Remove attachment"
                  ariaLabel="Remove attachment"
                  onClick={() => onRemoveExistingFile(file.FileName)}
                />
              )}
            </Stack>
          ))}

          {/* New Pending Attachments */}
          {newFiles.map((file, idx) => (
            <Stack
              key={`${file.name}-${idx}`}
              horizontal
              verticalAlign="center"
              horizontalAlign="space-between"
              style={{
                padding: '6px 12px',
                backgroundColor: '#eff6fc',
                borderRadius: 4,
                borderLeft: '3px solid #107c41'
              }}
            >
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Document" style={{ color: '#107c41' }} />
                <Text variant="medium">{file.name}</Text>
                <Text variant="small" style={{ color: '#666' }}>
                  ({formatFileSize(file.size)})
                </Text>
              </Stack>
              {!disabled && (
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  title="Remove attachment"
                  ariaLabel="Remove attachment"
                  onClick={() => handleRemoveNewFile(idx)}
                />
              )}
            </Stack>
          ))}
        </Stack>
      )}
    </Stack>
  );
};

export default FileUploadControl;
