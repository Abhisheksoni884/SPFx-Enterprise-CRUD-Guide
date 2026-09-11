import { BaseComponentContext } from '@microsoft/sp-component-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from './pnpjsConfig';
import { logError } from './LoggingService';

export interface INavigationNode {
  key: string;
  name: string;
  url?: string;
  icon?: string;
  isExpanded?: boolean;
  children?: INavigationNode[];
  nodeType: 'library' | 'list' | 'folder' | 'page';
}

/**
 * Fetches Site Hierarchy (Document Libraries, Folders, and Lists)
 */
export const getSiteHierarchy = async (context: BaseComponentContext | WebPartContext): Promise<INavigationNode[]> => {
  try {
    const sp = getSP(context);
    // Fetch user non-hidden lists & libraries
    const lists = await sp.web.lists
      .select('Id', 'Title', 'BaseTemplate', 'Hidden', 'DefaultViewUrl', 'RootFolder/ServerRelativeUrl')
      .expand('RootFolder')
      .filter('Hidden eq false')();

    const hierarchyNodes: INavigationNode[] = [];

    for (const list of lists) {
      // BaseTemplate 101 = Document Library, 119 = Site Pages (skip here to avoid duplication), 100 = Generic List
      if (list.BaseTemplate === 119) {
        continue; // SitePages will be handled separately in getSitePages()
      }

      const isDocLib = list.BaseTemplate === 101;
      const listNode: INavigationNode = {
        key: `list-${list.Id}`,
        name: list.Title,
        url: list.DefaultViewUrl || (list.RootFolder ? list.RootFolder.ServerRelativeUrl : '#'),
        icon: isDocLib ? 'FabricFolder' : 'List',
        nodeType: isDocLib ? 'library' : 'list',
        children: []
      };

      // If Document Library, fetch top-level subfolders
      if (isDocLib && list.RootFolder && list.RootFolder.ServerRelativeUrl) {
        try {
          const subFolders = await sp.web
            .getFolderByServerRelativePath(list.RootFolder.ServerRelativeUrl)
            .folders.select('Name', 'ServerRelativeUrl')
            .filter("Name ne 'Forms'")();

          listNode.children = subFolders.map(folder => ({
            key: `folder-${folder.ServerRelativeUrl}`,
            name: folder.Name,
            url: folder.ServerRelativeUrl,
            icon: 'FolderSearch',
            nodeType: 'folder' as const
          }));
        } catch (folderErr) {
          console.warn(`Could not fetch subfolders for list ${list.Title}`, folderErr);
        }
      }

      hierarchyNodes.push(listNode);
    }

    return hierarchyNodes;
  } catch (error) {
    logError(context as unknown as WebPartContext, 'getSiteHierarchy', 'SiteNavigationService', error).catch(() => undefined);
    throw error;
  }
};

/**
 * Fetches all Site Pages from the SitePages library
 */
export const getSitePages = async (context: BaseComponentContext | WebPartContext): Promise<INavigationNode[]> => {
  try {
    const sp = getSP(context);
    const sitePagesList = await sp.web.lists.getByTitle('Site Pages');

    const items = await sitePagesList.items
      .select('Id', 'Title', 'FileLeafRef', 'FileRef')
      .filter("FSObjType eq 0 and substringof('.aspx', FileLeafRef)")();

    // Filter out system/predefined pages (Header, Footer, Templates, Forms, etc.)
    const predefinedPatterns = [
      'header', 'footer', 'template', 'form', 'layout', 'search',
      'apppage', 'system', 'component'
    ];

    const pageMap = new Map<string, INavigationNode>();

    items.forEach(item => {
      const fileLeaf = (item.FileLeafRef || '').toLowerCase();
      const title = (item.Title || '').toLowerCase();

      // Check if page matches any predefined system pattern
      const isPredefined = predefinedPatterns.some(pattern => 
        fileLeaf.includes(pattern) || title.includes(pattern)
      );

      if (!isPredefined) {
        const pageTitle = item.Title || item.FileLeafRef.replace(/\.aspx$/i, '');
        const pageKey = pageTitle.trim().toLowerCase();

        // Deduplicate by title/page name
        if (!pageMap.has(pageKey)) {
          pageMap.set(pageKey, {
            key: `page-${item.Id}`,
            name: pageTitle,
            url: item.FileRef,
            icon: 'Page',
            nodeType: 'page' as const
          });
        }
      }
    });

    return Array.from(pageMap.values());
  } catch (error) {
    logError(context as unknown as WebPartContext, 'getSitePages', 'SiteNavigationService', error).catch(() => undefined);
    // Fallback query if 'Site Pages' localized title differs
    try {
      const sp = getSP(context);
      const pagesList = await sp.web.lists.ensure('SitePages');
      const items = await pagesList.list.items
        .select('Id', 'Title', 'FileLeafRef', 'FileRef')
        .filter("FSObjType eq 0 and substringof('.aspx', FileLeafRef)")();

      const predefinedPatterns = [
        'header', 'footer', 'template', 'form', 'layout', 'search',
        'apppage', 'system', 'component'
      ];

      const pageMap = new Map<string, INavigationNode>();

      items.forEach(item => {
        const fileLeaf = (item.FileLeafRef || '').toLowerCase();
        const title = (item.Title || '').toLowerCase();

        const isPredefined = predefinedPatterns.some(pattern => 
          fileLeaf.includes(pattern) || title.includes(pattern)
        );

        if (!isPredefined) {
          const pageTitle = item.Title || item.FileLeafRef.replace(/\.aspx$/i, '');
          const pageKey = pageTitle.trim().toLowerCase();

          if (!pageMap.has(pageKey)) {
            pageMap.set(pageKey, {
              key: `page-${item.Id}`,
              name: pageTitle,
              url: item.FileRef,
              icon: 'Page',
              nodeType: 'page' as const
            });
          }
        }
      });

      return Array.from(pageMap.values());
    } catch (fallbackErr) {
      logError(context as unknown as WebPartContext, 'getSitePages-fallback', 'SiteNavigationService', fallbackErr).catch(() => undefined);
      return [];
    }
  }
};

export default {
  getSiteHierarchy,
  getSitePages
};

