import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { ISiteSidebarProps } from './ISiteSidebarProps';
import { getSiteHierarchy, getSitePages, INavigationNode } from '../../../framework/Services/SiteNavigationService';
import { NAV_CONSTANTS } from '../../../framework/Constants/NavigationConstants';
import styles from './SiteSidebar.module.scss';

const SiteSidebar: React.FC<ISiteSidebarProps> = (props) => {
  const { context, isOpen: propsIsOpen, onDismiss } = props;

  const savedOpenState = sessionStorage.getItem('spfx_site_sidebar_open');
  const initialIsOpen = savedOpenState !== null ? savedOpenState === 'true' : propsIsOpen;

  const [isOpen, setIsOpenState] = useState<boolean>(initialIsOpen);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);
  const [hierarchyNodes, setHierarchyNodes] = useState<INavigationNode[]>([]);
  const [pageNodes, setPageNodes] = useState<INavigationNode[]>([]);
  const [isHierarchyExpanded, setIsHierarchyExpanded] = useState<boolean>(true);
  const [isPagesExpanded, setIsPagesExpanded] = useState<boolean>(true);
  const [expandedNodeKeys, setExpandedNodeKeys] = useState<Record<string, boolean>>({});

  const setSidebarOpen = useCallback((open: boolean): void => {
    sessionStorage.setItem('spfx_site_sidebar_open', String(open));
    setIsOpenState(open);
  }, []);

  const loadNavigationData = useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setErrorMessage(undefined);
    try {
      const [hierarchy, pages] = await Promise.all([
        getSiteHierarchy(context),
        getSitePages(context)
      ]);

      setHierarchyNodes(hierarchy);
      setPageNodes(pages);
      setIsLoading(false);
    } catch {
      setIsLoading(false);
      setErrorMessage(NAV_CONSTANTS.ERROR_TEXT);
    }
  }, [context]);

  useEffect(() => {
    loadNavigationData().catch((err) => console.error(err));
  }, [loadNavigationData]);

  useEffect(() => {
    setIsOpenState(propsIsOpen);
    if (propsIsOpen && hierarchyNodes.length === 0) {
      loadNavigationData().catch((err) => console.error(err));
    }
  }, [propsIsOpen, hierarchyNodes.length, loadNavigationData]);

  const toggleGroup = (group: 'hierarchy' | 'pages'): void => {
    if (group === 'hierarchy') {
      setIsHierarchyExpanded((prev) => !prev);
    } else {
      setIsPagesExpanded((prev) => !prev);
    }
  };

  const toggleNodeExpand = (key: string, e: React.MouseEvent): void => {
    e.stopPropagation();
    e.preventDefault();
    setExpandedNodeKeys((prev) => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  const navigateToUrl = (url?: string): void => {
    if (url && url !== '#') {
      sessionStorage.setItem('spfx_site_sidebar_open', 'true');

      try {
        const currentUrlParams = new URLSearchParams(window.location.search);
        const debugManifests = currentUrlParams.get('debugManifestsFile');
        const customActions = currentUrlParams.get('customActions');
        const loadSPFX = currentUrlParams.get('loadSPFX');

        if (debugManifests || customActions) {
          const targetUrlObj = new URL(url, window.location.origin);
          if (debugManifests) targetUrlObj.searchParams.set('debugManifestsFile', debugManifests);
          if (customActions) targetUrlObj.searchParams.set('customActions', customActions);
          if (loadSPFX) targetUrlObj.searchParams.set('loadSPFX', loadSPFX);

          window.location.href = targetUrlObj.toString();
          return;
        }
      } catch {
        // Fallback standard navigation
      }

      window.location.href = url;
    }
  };

  const renderNode = (node: INavigationNode, depth: number = 0): React.ReactNode => {
    const hasChildren = node.children && node.children.length > 0;
    const isExpanded = !!expandedNodeKeys[node.key];

    return (
      <div key={node.key}>
        <div
          className={styles.nodeItem}
          onClick={() => navigateToUrl(node.url)}
          title={node.name}
        >
          <div className={styles.itemLeft}>
            {hasChildren ? (
              <Icon
                iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'}
                className={styles.itemIcon}
                onClick={(e) => toggleNodeExpand(node.key, e)}
              />
            ) : (
              <Icon iconName={node.icon || 'Page'} className={styles.itemIcon} />
            )}
            <span className={styles.itemText}>{node.name}</span>
          </div>

          {node.url && (
            <Icon iconName="NavigateExternalInline" className={styles.itemIcon} />
          )}
        </div>

        {hasChildren && isExpanded && (
          <div className={styles.childContainer}>
            {node.children!.map((child) => renderNode(child, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={styles.siteSidebarContainer}>
      {/* Fixed Arrow Toggle Tab */}
      {!isOpen && (
        <button
          className={styles.toggleLauncherBtn}
          onClick={() => setSidebarOpen(true)}
          title="Open Site Navigation"
        >
          <Icon iconName="ChevronRight" className={styles.btnIcon} />
        </button>
      )}

      {/* Slide-in Fluent UI Navigation Panel */}
      <Panel
        isOpen={isOpen}
        onDismiss={() => {
          setSidebarOpen(false);
          if (onDismiss) {
            onDismiss();
          }
        }}
        type={PanelType.customNear}
        customWidth="320px"
        isLightDismiss={true}
        hasCloseButton={false}
        headerText=""
        onRenderHeader={() => (
          <div className={styles.panelHeader}>
            <div className={styles.headerLeftTitle}>
              <Icon iconName="CompassNW" className={styles.headerIcon} />
              <h2 className={styles.headerTitle}>{NAV_CONSTANTS.TITLE}</h2>
            </div>
            <button
              className={styles.closeBtn}
              onClick={() => setSidebarOpen(false)}
              title="Close Navigation"
            >
              <Icon iconName="Cancel" />
            </button>
          </div>
        )}
      >
        <div className={styles.panelBody}>
          {isLoading ? (
            <div className={styles.loadingContainer}>
              <Spinner size={SpinnerSize.large} label={NAV_CONSTANTS.LOADING_TEXT} />
            </div>
          ) : (
            <>
              {errorMessage && (
                <div className={styles.errorBanner}>{errorMessage}</div>
              )}

              {/* 1. SITE HIERARCHY BRANCH */}
              <div className={styles.navGroup}>
                <div
                  className={styles.groupHeader}
                  onClick={() => toggleGroup('hierarchy')}
                >
                  <div className={styles.headerLeft}>
                    <Icon iconName={isHierarchyExpanded ? 'ChevronDown' : 'ChevronRight'} />
                    <Icon iconName="Org" />
                    <span>{NAV_CONSTANTS.HIERARCHY_PARENT_LABEL}</span>
                  </div>
                  <span className={styles.counterBadge}>{hierarchyNodes.length}</span>
                </div>

                {isHierarchyExpanded && (
                  <div className={styles.groupContent}>
                    {hierarchyNodes.length === 0 ? (
                      <div className={styles.nodeItem}>{NAV_CONSTANTS.EMPTY_HIERARCHY}</div>
                    ) : (
                      hierarchyNodes.map((node) => renderNode(node))
                    )}
                  </div>
                )}
              </div>

              {/* 2. ALL SITE PAGES BRANCH */}
              <div className={styles.navGroup}>
                <div
                  className={styles.groupHeader}
                  onClick={() => toggleGroup('pages')}
                >
                  <div className={styles.headerLeft}>
                    <Icon iconName={isPagesExpanded ? 'ChevronDown' : 'ChevronRight'} />
                    <Icon iconName="DocumentSet" />
                    <span>{NAV_CONSTANTS.PAGES_PARENT_LABEL}</span>
                  </div>
                  <span className={styles.counterBadge}>{pageNodes.length}</span>
                </div>

                {isPagesExpanded && (
                  <div className={styles.groupContent}>
                    {pageNodes.length === 0 ? (
                      <div className={styles.nodeItem}>{NAV_CONSTANTS.EMPTY_PAGES}</div>
                    ) : (
                      pageNodes.map((node) => renderNode(node))
                    )}
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </Panel>
    </div>
  );
};

export default SiteSidebar;

