import * as React from 'react';
import { useState } from 'react';
import { NormalPeoplePicker, IPersonaProps } from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './CrudOperation.module.scss';
import { IPersonaInfo } from '../models/IEmployeeLeaveRequest';

export interface IPeoplePickerFieldProps {
  context: WebPartContext;
  label: string;
  required?: boolean;
  errorMessage?: string;
  disabled?: boolean;
  selectedPerson?: IPersonaInfo;
  onChange: (person: IPersonaInfo | undefined) => void;
}

// Shape of a single suggestion returned by the SharePoint People Picker REST endpoint.
interface IClientPeoplePickerResult {
  Key: string;
  DisplayText: string;
  EntityData?: { Email?: string };
}

function toPersona(info: IPersonaInfo): IPersonaProps {
  return { key: info.loginName, text: info.displayName, secondaryText: info.email };
}

// Thin wrapper around Fluent UI's NormalPeoplePicker that resolves suggestions
// directly from SharePoint's ClientPeoplePickerWebServiceInterface REST endpoint,
// avoiding the need for a heavier third-party people-picker package.
const PeoplePickerField: React.FC<IPeoplePickerFieldProps> = (props) => {
  const { context, label, required, errorMessage, disabled, selectedPerson, onChange } = props;
  const [selectedItems, setSelectedItems] = useState<IPersonaProps[]>(
    selectedPerson ? [toPersona(selectedPerson)] : []
  );

  const onResolveSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
    if (!filterText || filterText.length < 2) {
      return [];
    }

    const requestBody = JSON.stringify({
      queryParams: {
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: 15,
        PrincipalSource: 15, // All sources (UserInfoList + Search + custom providers)
        PrincipalType: 1, // 1 = User
        QueryString: filterText
      }
    });

    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const response: SPHttpClientResponse = await context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
      body: requestBody
    });

    const json = await response.json();
    const rawResults = json.value ?? (json.d && json.d.ClientPeoplePickerSearchUser) ?? '[]';
    const results: IClientPeoplePickerResult[] = JSON.parse(rawResults);

    return results.map((result) => ({
      key: result.Key,
      text: result.DisplayText,
      secondaryText: (result.EntityData && result.EntityData.Email) || result.Key
    }));
  };

  const onChangePicker = (items?: IPersonaProps[]): void => {
    const picked = items && items[0];
    setSelectedItems(items || []);

    if (!picked) {
      onChange(undefined);
      return;
    }

    onChange({
      id: 0, // resolved to a real SharePoint user Id later, via LeaveRequestService.resolveEmployee
      loginName: picked.key as string,
      displayName: picked.text || '',
      email: picked.secondaryText || ''
    });
  };

  return (
    <div className={styles.peoplePickerField}>
      <span className={styles.fieldLabel}>
        {label}
        {required ? ' *' : ''}
      </span>
      <NormalPeoplePicker
        onResolveSuggestions={onResolveSuggestions}
        getTextFromItem={(persona: IPersonaProps) => persona.text || ''}
        pickerSuggestionsProps={{ suggestionsHeaderText: 'Suggested people', noResultsFoundText: 'No results found' }}
        selectedItems={selectedItems}
        onChange={onChangePicker}
        itemLimit={1}
        disabled={disabled}
        inputProps={{ placeholder: 'Type a name or email', 'aria-label': label }}
      />
      {errorMessage && <span className={styles.fieldError}>{errorMessage}</span>}
    </div>
  );
};

export default PeoplePickerField;
