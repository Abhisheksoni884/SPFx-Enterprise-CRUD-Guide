import { WebPartContext } from '@microsoft/sp-webpart-base';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { getSP } from './pnpjsConfig';
import { LIST_NAMES, TEAM_MAPPING_SELECT_FIELDS } from '../Constants/Constant';
import { logError } from './LoggingService';
import { IPersonaInfo } from '../../webparts/crudOperation/models/IEmployeeLeaveRequest';

interface IPersonRaw {
  Id: number;
  Title?: string;
  EMail?: string;
  Name?: string;
}

interface ITeamMappingRawItem {
  Id: number;
  Manager?: IPersonRaw;
  Member?: IPersonRaw | IPersonRaw[];
}

function matchesPerson(
  person?: IPersonRaw,
  userEmail?: string,
  userLogin?: string,
  userName?: string
): boolean {
  if (!person) {
    return false;
  }

  const targets = [
    userEmail ? userEmail.toLowerCase().trim() : '',
    userLogin ? userLogin.toLowerCase().trim() : '',
    userName ? userName.toLowerCase().trim() : ''
  ].filter((t) => t.length > 0);

  if (targets.length === 0) {
    return false;
  }

  const personValues = [
    person.EMail ? person.EMail.toLowerCase().trim() : '',
    person.Name ? person.Name.toLowerCase().trim() : '',
    person.Title ? person.Title.toLowerCase().trim() : ''
  ].filter((v) => v.length > 0);

  return targets.some((t) =>
    personValues.some((v) => v === t || (t.indexOf('@') !== -1 && v.indexOf(t) !== -1))
  );
}

/**
 * Retrieves all records from the Manager_List / Team Mapping list with fallback list name resolution.
 */
export const getAllTeamMappings = async (
  context: BaseComponentContext | WebPartContext,
  listName: string = LIST_NAMES.TEAM_MAPPING
): Promise<ITeamMappingRawItem[]> => {
  const sp = getSP(context);
  const candidateListNames = [listName, 'Manager_List', 'Manager List', 'Team Mapping'];

  for (const targetName of candidateListNames) {
    try {
      return await sp.web.lists.getByTitle(targetName).items
        .select(...TEAM_MAPPING_SELECT_FIELDS)
        .expand('Manager', 'Member')();
    } catch {
      try {
        return await sp.web.lists.getByTitle(targetName).items
          .expand('Manager', 'Member')();
      } catch {
        // Continue trying next candidate name if list not found under current name
      }
    }
  }

  await logError(context, 'getAllTeamMappings', 'TeamMappingService', 'Failed to find Manager_List or Team Mapping list');
  return [];
};

/**
 * Finds the Manager assigned to a specific employee.
 */
export const getManagerForEmployee = async (
  context: BaseComponentContext | WebPartContext,
  employeeEmail?: string,
  employeeLogin?: string,
  employeeName?: string,
  listName: string = LIST_NAMES.TEAM_MAPPING
): Promise<IPersonaInfo | undefined> => {
  const mappings = await getAllTeamMappings(context, listName);

  for (const item of mappings) {
    if (!item.Manager || !item.Member) {
      continue;
    }

    const members = Array.isArray(item.Member) ? item.Member : [item.Member];
    const isMemberMatch = members.some((m) => matchesPerson(m, employeeEmail, employeeLogin, employeeName));

    if (isMemberMatch && item.Manager) {
      return {
        id: item.Manager.Id,
        loginName: item.Manager.Name || item.Manager.EMail || item.Manager.Title || '',
        displayName: item.Manager.Title || '',
        email: item.Manager.EMail || ''
      };
    }
  }

  return undefined;
};

/**
 * Gets a list of employee identifiers (emails, titles, ids) reporting to a manager.
 */
export const getManagedTeamMembers = async (
  context: BaseComponentContext | WebPartContext,
  managerEmail?: string,
  managerLogin?: string,
  managerName?: string,
  listName: string = LIST_NAMES.TEAM_MAPPING
): Promise<{ emails: string[]; titles: string[]; ids: number[] }> => {
  const mappings = await getAllTeamMappings(context, listName);
  const emails = new Set<string>();
  const titles = new Set<string>();
  const ids = new Set<number>();

  for (const item of mappings) {
    if (matchesPerson(item.Manager, managerEmail, managerLogin, managerName)) {
      if (item.Member) {
        const members = Array.isArray(item.Member) ? item.Member : [item.Member];
        members.forEach((m) => {
          if (m.EMail) emails.add(m.EMail.toLowerCase().trim());
          if (m.Title) titles.add(m.Title.toLowerCase().trim());
          if (m.Id) ids.add(m.Id);
        });
      }
    }
  }

  return {
    emails: Array.from(emails),
    titles: Array.from(titles),
    ids: Array.from(ids)
  };
};

/**
 * Checks if the user is listed as a Manager in the Manager_List list.
 */
export const isUserAManager = async (
  context: BaseComponentContext | WebPartContext,
  userEmail?: string,
  userLogin?: string,
  userName?: string,
  listName: string = LIST_NAMES.TEAM_MAPPING
): Promise<boolean> => {
  const mappings = await getAllTeamMappings(context, listName);
  return mappings.some((item) => matchesPerson(item.Manager, userEmail, userLogin, userName));
};

export default {
  getAllTeamMappings,
  getManagerForEmployee,
  getManagedTeamMembers,
  isUserAManager
};

