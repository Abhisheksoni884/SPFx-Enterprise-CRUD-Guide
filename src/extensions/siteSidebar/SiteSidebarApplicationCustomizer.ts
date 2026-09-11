import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import SiteSidebar from './components/SiteSidebar';
import { ISiteSidebarProps } from './components/ISiteSidebarProps';

const LOG_SOURCE: string = 'SiteSidebarApplicationCustomizer';

export interface ISiteSidebarApplicationCustomizerProperties {
  title?: string;
}

export default class SiteSidebarApplicationCustomizer
  extends BaseApplicationCustomizer<ISiteSidebarApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _hostContainer: HTMLElement | undefined;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initializing SiteSidebarApplicationCustomizer`);

    // Hide standard SharePoint headers, top suitebar, site navigation, and page title banner
    this._injectCleanViewCSS();

    // Listen to placeholder changes (handles SPA routing smoothly)
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);

    // Initial render call
    this._renderPlaceholders();

    return Promise.resolve();
  }

  private _injectCleanViewCSS(): void {
    const styleId = 'spfx-clean-view-style';
    if (!document.getElementById(styleId)) {
      const style = document.createElement('style');
      style.id = styleId;
      style.textContent = `
        /* Hide Left SharePoint App Bar (Discover, Publish, Build, OneDrive) */
        #sp-appBar, div[data-automation-id="sp-appBar"], [class*="appBar"] {
          display: none !important;
          width: 0 !important;
        }

        /* Shift main page content to full left edge */
        #spPageCanvasContent, #spCommandBar, #workbenchPageContent, #sp-theme-container, main {
          margin-left: 0 !important;
          left: 0 !important;
        }

        /* Hide O365 Top Bar / Suite Nav */
        #O365_MainLink_NavMenu, #suiteBarTop, #SuiteNavWrapper, #spCommandBar {
          display: none !important;
        }

        /* Hide Site Header & Horizontal Navigation */
        #spSiteHeader, div[data-automation-id="SiteHeader"], div[role="banner"], [class*="siteHeader"] {
          display: none !important;
        }

        /* Hide Page Title Banner / Hero Image Banner */
        div[data-automation-id="pageHeader"], div[class*="pageTitle"], [class*="titleBanner"] {
          display: none !important;
        }

        /* Hide Page Footer & Social Comments Section */
        #sp-footer, footer, div[id*="footer"], [class*="spFooter"], [class*="pageFooter"] {
          display: none !important;
        }

        /* Hide Page Likes, Comments, Views & Action Bar */
        div[class*="socialBar"], div[data-automation-id="socialBar"], div[class*="commentsWrapper"], #commentsHeader, [class*="commentsContainer"] {
          display: none !important;
        }

        /* Ensure main canvas takes full focus */
        #spPageCanvasContent {
          padding-top: 10px !important;
        }
      `;
      document.head.appendChild(style);
    }
  }

  private _renderPlaceholders = (): void => {
    // 1. Try PlaceholderName.Top first
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
    }

    let targetElement: HTMLElement | null = null;

    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      targetElement = this._topPlaceholder.domElement;
    } else {
      // 2. Persistent DOM Fallback: Attach directly to document.body so the sidebar is never removed
      if (!this._hostContainer) {
        this._hostContainer = document.getElementById('spfx-site-sidebar-root') as HTMLElement;
        if (!this._hostContainer) {
          this._hostContainer = document.createElement('div');
          this._hostContainer.id = 'spfx-site-sidebar-root';
          document.body.appendChild(this._hostContainer);
        }
      }
      targetElement = this._hostContainer;
    }

    if (targetElement) {
      const isInitiallyOpen = sessionStorage.getItem('spfx_site_sidebar_open') === 'true';

      const element: React.ReactElement<ISiteSidebarProps> = React.createElement(
        SiteSidebar,
        {
          context: this.context,
          isOpen: isInitiallyOpen,
          onDismiss: () => {
            sessionStorage.setItem('spfx_site_sidebar_open', 'false');
          }
        }
      );

      ReactDom.render(element, targetElement);
    }
  };

  private _onDispose = (): void => {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this._topPlaceholder.domElement);
    }
    if (this._hostContainer) {
      ReactDom.unmountComponentAtNode(this._hostContainer);
    }
  };
}
