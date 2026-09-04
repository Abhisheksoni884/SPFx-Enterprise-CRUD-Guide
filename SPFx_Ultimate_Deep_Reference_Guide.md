# SharePoint Framework (SPFx) Reference Guide

**Last Updated**: September 2026  
**Latest Stable**: SPFx 1.23.x | **Preview**: SPFx 1.24 (React 18 + Copilot Components)  
**Scope**: SharePoint Online, Microsoft Teams, Viva Connections, Microsoft 365 Copilot

## Table of Contents

1. [What Is SPFx?](#1-what-is-spfx)
2. [Component Types](#2-component-types)
3. [Client-Side Web Parts](#3-client-side-web-parts)
4. [SPFx Extensions](#4-spfx-extensions)
5. [Adaptive Card Extensions (ACEs)](#5-adaptive-card-extensions-aces)
6. [Library Components](#6-library-components)
7. [Toolchain & Build Pipeline](#7-toolchain--build-pipeline)
8. [Security Architecture](#8-security-architecture)
9. [Microsoft Graph Integration](#9-microsoft-graph-integration)
10. [Deployment Models](#10-deployment-models)
11. [Microsoft 365 Platform Integrations](#11-microsoft-365-platform-integrations)
12. [SPFx Version History](#12-spfx-version-history)
13. [Performance Optimization](#13-performance-optimization)
14. [Developer Tooling](#14-developer-tooling)
15. [PnP Community & Resources](#15-pnp-community--resources)
16. [Advanced Topics](#16-advanced-topics)

## 1. What Is SPFx?

The SharePoint Framework (SPFx) is the official development model for extending Microsoft 365. It provides a client-side approach for building solutions that integrate with SharePoint Online, Microsoft Teams, Viva Connections, and Microsoft 365 Copilot.

### Core Principles

- **Client-Side Only** - All code runs in browser with OAuth 2.0 authentication
- **Modern Web Standards** - TypeScript, React, npm ecosystem
- **Evergreen Compatibility** - Survives Microsoft 365 platform updates
- **Framework Agnostic** - React (recommended), Angular, Vue, or Vanilla JS
- **Unified Extensibility** - One codebase across SharePoint, Teams, Viva, Copilot

### What SPFx Replaced

| Legacy Model | SPFx Replacement |
|---|---|
| SharePoint Add-In Model | Web Parts + Extensions |
| Farm Solutions | SPFx Solutions |
| JSOM/CSOM Scripts | PnPjs + MSGraphClient |
| Script Editor Web Part | SPFx Web Parts |
| Domain Isolated Web Parts | Standard Web Parts + AAD Auth |

## 2. Component Types

**SPFx Component Hierarchy:**
- **Web Parts** - Configurable UI blocks for pages
- **Extensions**
  - Application Customizers (page headers/footers)
  - Field Customizers (column rendering)
  - ListView Command Sets (toolbar/menu actions)
  - Form Customizers (form overrides)
  - Search Query Modifiers (query interceptors)
- **Adaptive Card Extensions (ACEs)** - Viva Connections cards
- **Library Components** - Shared code libraries

| Component Type | Renders In | Configurable | Min Version |
|---|---|---|---|
| Web Part | Pages, Teams, Viva | Yes | v1.0 |
| Application Customizer | All modern pages | No | v1.0 |
| Field Customizer | List columns | No | v1.0 |
| ListView Command Set | List toolbar/menu | No | v1.0 |
| Form Customizer | New/Edit/View forms | No | v1.15 |
| Search Query Modifier | Search pipeline | No | v1.16 |
| ACE | Viva Dashboard | Yes | v1.13 |
| Library Component | No UI | N/A | v1.9.1 |

## 3. Client-Side Web Parts

### Web Part Structure

```
src/webparts/myWebPart/
├── MyWebPartWebPart.ts              # Main web part class
├── MyWebPartWebPart.manifest.json   # Metadata & registration
├── components/
│   ├── MyWebPart.tsx                # React component
│   ├── MyWebPart.module.scss        # Scoped CSS
│   └── IMyWebPartProps.ts           # Props interface
└── loc/
    ├── en-us.js                     # Localization strings
    └── mystrings.d.ts               # Type definitions
```

### Essential Web Part Class

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

export default class MyWebPartWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  
  public render(): void {
    const element = React.createElement(MyWebPart, {
      context: this.context,
      properties: this.properties
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupFields: [
            PropertyPaneTextField('listName', { label: 'List Name' })
          ]
        }]
      }]
    };
  }
}
```

### Key Context Properties

```typescript
// Page & Site info
this.context.pageContext.web.absoluteUrl
this.context.pageContext.user.displayName
this.context.pageContext.user.email

// HTTP Clients
this.context.spHttpClient          // SharePoint REST API
this.context.httpClient            // Generic HTTP
this.context.msGraphClientFactory  // Microsoft Graph
this.context.aadHttpClientFactory  // Azure AD secured HTTP

// Other APIs
this.context.sdks.microsoftTeams   // Teams SDK (when in Teams)
this.context.propertyPane          // Property pane operations
this.context.dynamicDataProvider   // Dynamic data connections
```

### Supported Frameworks

| Framework | Support Level |
|---|---|
| React | Primary/Recommended |
| No Framework | Supported |
| Angular | Community Support |
| Vue.js | Community Support |

## 4. SPFx Extensions

### Application Customizers

Inject custom HTML/scripts into page headers/footers.

```typescript
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

export default class GlobalNavApplicationCustomizer
  extends BaseApplicationCustomizer<{}> {

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders);
    this._renderPlaceholders();
    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    const topPlaceholder = this.context.placeholderProvider
      .tryCreateContent(PlaceholderName.Top);

    if (topPlaceholder?.domElement) {
      topPlaceholder.domElement.innerHTML = `<div>Custom Header</div>`;
    }
  }
}
```

**Register via PnP PowerShell:**
```powershell
Add-PnPCustomAction -Title "GlobalNav" `
  -Location "ClientSideExtension.ApplicationCustomizer" `
  -ClientSideComponentId "your-guid" `
  -Scope Tenant
```

### Field Customizers

Override column rendering in lists/libraries.

```typescript
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

export default class StatusBadgeFieldCustomizer
  extends BaseFieldCustomizer<{}> {

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const value = event.fieldValue || 'Unknown';
    const colors = {
      'Active': '#107c10',
      'Pending': '#f7630c',
      'Closed': '#d13438'
    };
    
    event.domElement.innerHTML = `
      <span style="background:${colors[value]};color:white;padding:2px 8px;border-radius:12px;">
        ${value}
      </span>`;
  }
}
```

### ListView Command Sets

Add custom buttons to list toolbars and context menus.

```typescript
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

export default class ApproveCommandSet extends BaseListViewCommandSet<{}> {

  public onListViewUpdated(event): void {
    const approveCmd = this.tryGetCommand('COMMAND_APPROVE');
    if (approveCmd) {
      approveCmd.visible = event.selectedRows.length === 1;
    }
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_APPROVE':
        const row = event.selectedRows[0];
        const itemId = row.getValueByName('ID');
        await this._approveItem(itemId);
        break;
    }
  }
}
```

### Form Customizers

Replace default New/Edit/View forms (SPFx v1.15+).

```typescript
import { BaseFormCustomizer } from '@microsoft/sp-listview-extensibility';

export default class TaskFormCustomizer extends BaseFormCustomizer<{}> {

  public render(): void {
    const element = React.createElement(CustomForm, {
      displayMode: this.displayMode,  // 'New' | 'Edit' | 'Display'
      item: this.item,
      onSave: this.onSave.bind(this),
      onClose: this.onClose.bind(this)
    });
    ReactDOM.render(element, this.domElement);
  }
}
```

### Search Query Modifiers

Intercept and modify search queries (SPFx v1.16+).

```typescript
import { BaseSearchQueryModifier } from '@microsoft/sp-search-extensibility';

export default class DepartmentFilterSearchModifier
  extends BaseSearchQueryModifier<{}> {

  public async modifyQuery(input): Promise<any> {
    const userDept = this.context.pageContext.user.department;
    return {
      queryText: `${input.queryText} AND Department:"${userDept}"`
    };
  }
}
```

## 5. Adaptive Card Extensions (ACEs)

ACEs power Microsoft Viva Connections Dashboard cards.

### ACE Structure

- **Card View** - Compact dashboard view (Basic, PrimaryText, Image, TextInput, Chart types)
- **Quick View** - Full modal dialog using Adaptive Card JSON templates

### ACE Main Class

```typescript
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';

export default class MyAce extends BaseAdaptiveCardExtension<IProps, IState> {

  public async onInit(): Promise<void> {
    this.state = {
      pendingCount: 0,
      items: []
    };

    this.cardNavigator.register(CARD_VIEW_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_ID, () => new QuickView());

    await this._loadData();
    return Promise.resolve();
  }

  private async _loadData(): Promise<void> {
    const client = await this.context.msGraphClientFactory.getClient('3');
    const result = await client.api('/me/pendingAccessReviewInstances').get();
    this.setState({ pendingCount: result.value.length, items: result.value });
  }
}
```

### Card View

```typescript
import { BaseBasicCardView } from '@microsoft/sp-adaptive-card-extension-base';

export class CardView extends BaseBasicCardView<IProps, IState> {

  public get cardButtons() {
    return [{
      title: 'View All',
      action: {
        type: 'QuickView',
        parameters: { view: QUICK_VIEW_ID }
      }
    }];
  }

  public get data() {
    return {
      primaryText: `${this.state.pendingCount} Pending Approvals`,
      title: this.properties.title
    };
  }
}
```

### Quick View

```typescript
import { BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';

export class QuickView extends BaseAdaptiveCardView<IProps, IState, IData> {
  
  public get data() {
    return {
      items: this.state.items,
      totalCount: this.state.pendingCount
    };
  }

  public get template() {
    return require('./template/QuickViewTemplate.json');
  }

  public async onAction(action): Promise<void> {
    if (action.type === 'Submit') {
      await this._approveItem(action.data.id);
    }
  }
}
```

## 6. Library Components

Shared code libraries that multiple SPFx solutions can consume.

### Key Characteristics

- No UI
- Only ONE version per tenant
- Must be deployed to Tenant App Catalog
- Prevents code duplication

### Creating a Library

```typescript
// src/index.ts (public API)
export { SharedUtilities } from './SharedUtilities';
export { GraphService } from './services/GraphService';
export { IUserProfile } from './models/IUserProfile';
```

### Consuming a Library

```typescript
// In consuming web part's config/config.json
{
  "externals": {
    "my-spfx-library": {
      "path": "https://tenant.sharepoint.com/sites/appcatalog/.../my-library.js",
      "globalName": "my_spfx_library"
    }
  }
}

// Usage in code
import { SharedUtilities } from 'my-spfx-library';
const formatted = SharedUtilities.formatDate(new Date());
```

## 7. Toolchain & Build Pipeline

### Gulp-Based (SPFx <= 1.21.x)

```bash
gulp serve                      # Start local dev
gulp serve --nobrowser          # Serve without browser
gulp build                      # Development build
gulp bundle --ship              # Production bundle (minified)
gulp package-solution --ship    # Create .sppkg (production)
gulp clean                      # Clean build output
gulp trust-dev-cert            # Trust dev certificate
```

### Heft-Based (SPFx >= 1.22.x)

SPFx 1.22+ uses Heft (Rush Stack) instead of Gulp.

```bash
heft start                      # Start local dev
heft build --production         # Production build
heft test                       # Run tests
heft clean                      # Clean output
```

### Node.js Compatibility

| SPFx Version | Node.js | Toolchain | TypeScript |
|---|---|---|---|
| 1.23-1.24 | 22 LTS | Heft | 5.8 |
| 1.22 | 22 LTS | Heft | 5.8 |
| 1.21 | 18 or 22 | Gulp | 4.x |
| 1.18-1.20 | 16-18 | Gulp | 4.x |

**Use NVM for version management:**
```bash
nvm install 22
nvm use 22
```

### Key Config Files

**config/package-solution.json:**
```json
{
  "solution": {
    "name": "my-solution",
    "id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "skipFeatureDeployment": true,
    "webApiPermissionRequests": [
      { "resource": "Microsoft Graph", "scope": "User.Read" }
    ]
  }
}
```

## 8. Security Architecture

### Content Security Policy (CSP)

**Enforced since March 2026** - SPFx must be CSP-compliant.

**CSP Blocks:**
- Inline JavaScript
- `eval()` and dynamic code execution
- Untrusted external scripts

**Making Solutions CSP-Compliant:**

```typescript
// WRONG
this.domElement.innerHTML = '<script>alert("hello")</script>';

// CORRECT
const button = document.createElement('button');
button.addEventListener('click', () => alert('hello'));
this.domElement.appendChild(button);
```

### API Permissions

Declare permissions in `package-solution.json`:

```json
"webApiPermissionRequests": [
  { "resource": "Microsoft Graph", "scope": "User.Read" },
  { "resource": "Microsoft Graph", "scope": "Sites.Read.All" },
  { "resource": "ContosoAPI", "scope": "user_impersonation" }
]
```

Admin must approve in **SharePoint Admin Center → Advanced → API access**.

### Security Best Practices

- Request minimum required API scopes
- Never hardcode secrets - use Azure Key Vault
- Use MSGraphClientV3 for automatic token management
- Validate and sanitize user input
- Run `npm audit` regularly
- HTTPS only for external requests

## 9. Microsoft Graph Integration

### MSGraphClientV3

Recommended way to call Microsoft Graph - handles OAuth automatically.

```typescript
import { MSGraphClientV3 } from '@microsoft/sp-http';

export default class MyWebPart extends BaseClientSideWebPart<{}> {

  public async render(): Promise<void> {
    const client: MSGraphClientV3 =
      await this.context.msGraphClientFactory.getClient('3');

    // GET current user
    const me = await client.api('/me').get();

    // GET with filtering and sorting
    const messages = await client
      .api('/me/messages')
      .select('subject,from,receivedDateTime')
      .top(10)
      .orderby('receivedDateTime desc')
      .get();

    // POST
    await client.api('/teams/{team-id}/channels/{channel-id}/messages')
      .post({
        body: {
          content: 'Hello from SPFx!',
          contentType: 'text'
        }
      });

    // Handle pagination
    let allItems = [];
    let response = await client.api('/me/drive/root/children').get();
    allItems = [...response.value];
    while (response['@odata.nextLink']) {
      response = await client.api(response['@odata.nextLink']).get();
      allItems = [...allItems, ...response.value];
    }
  }
}
```

### PnPjs Integration

Community-maintained fluent API for SharePoint and Graph.

```bash
npm install @pnp/sp @pnp/graph @pnp/logging --save
```

```typescript
import { spfi, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class MyWebPart extends BaseClientSideWebPart<{}> {

  private sp;
  private graph;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(spSPFx(this.context));
    this.graph = graphfi().using(graphSPFx(this.context));
  }

  public async render(): Promise<void> {
    // SharePoint operations
    const web = await this.sp.web();
    const lists = await this.sp.web.lists.filter('Hidden eq false')();
    
    // Get items with expand
    const items = await this.sp.web.lists
      .getByTitle('Tasks')
      .items
      .select('Title', 'Status', 'AssignedTo/Title')
      .expand('AssignedTo')
      .top(50)();

    // Create item
    await this.sp.web.lists.getByTitle('Tasks').items.add({
      Title: 'New Task',
      Status: 'Active'
    });

    // Graph operations
    const me = await this.graph.me();
  }
}
```

## 10. Deployment Models

### Tenant App Catalog

Central repository for deploying SPFx across the entire tenant.

```powershell
# PnP PowerShell
Connect-PnPOnline -Url "https://tenant-admin.sharepoint.com" -Interactive

Add-PnPApp -Path ".\solution\my-solution.sppkg" -Scope Tenant -Publish -SkipFeatureDeployment
```

```bash
# CLI for Microsoft 365
m365 spo app add --filePath ./solution/my-solution.sppkg --scope tenant --overwrite
m365 spo app deploy --name my-solution-client-side-solution --scope tenant
```

### Site Collection App Catalog

Deploy to specific site without tenant admin rights.

```powershell
# Enable site collection app catalog (once, by tenant admin)
Add-PnPSiteCollectionAppCatalog -Site "https://tenant.sharepoint.com/sites/mysite"

# Deploy to site catalog (site admin can do this)
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/mysite" -Interactive
Add-PnPApp -Path ".\solution\my-solution.sppkg" -Scope Site -Publish
```

### Tenant-Wide Deployment

Set `skipFeatureDeployment: true` to auto-deploy to all sites.

```json
{
  "solution": {
    "skipFeatureDeployment": true
  }
}
```

### Deployment Scope Comparison

| Scope | Who Deploys | Who Can Use | Use Case |
|---|---|---|---|
| Tenant App Catalog | Tenant Admin | All sites | Organization-wide |
| Site Collection Catalog | Site Admin | That site only | Department-specific |
| Tenant-Wide | Tenant Admin | All sites auto | Global branding/nav |

## 11. Microsoft 365 Platform Integrations

### Microsoft Teams Integration

Enable Teams in manifest:

```json
{
  "supportedHosts": [
    "SharePointWebPart",
    "TeamsPersonalApp",
    "TeamsTab"
  ]
}
```

Detect and use Teams context:

```typescript
const isTeams = !!this.context.sdks.microsoftTeams;

if (isTeams) {
  const teamsContext = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
  const teamId = teamsContext.team?.internalId;
  const theme = teamsContext.app.theme;

  // Listen to theme changes
  this.context.sdks.microsoftTeams.teamsJs.app.registerOnThemeChangeHandler((theme) => {
    this.setState({ teamsTheme: theme });
  });
}
```

### Viva Connections

Viva Connections = Employee Experience platform built on SharePoint + Teams.

**Architecture:**
- Dashboard → ACEs (Adaptive Card Extensions)
- Feed → SharePoint news + Viva Engage
- Resources → SharePoint pages and links

ACEs are the ONLY way to customize Viva Dashboard.

### Microsoft 365 Copilot

**New in SPFx 1.24 (Preview)**: Render interactive UI directly in Copilot conversations.

```typescript
import { BaseCopilotComponent } from '@microsoft/sp-copilot-base';

export default class ApprovalsComponent extends BaseCopilotComponent<IProps, IState> {

  public render(): React.ReactElement {
    return (
      <div className="copilot-component">
        <h3>Pending Approvals ({this.state.items.length})</h3>
        {this.state.items.map(item => (
          <div key={item.id}>
            <strong>{item.title}</strong>
            <button onClick={() => this._handleApprove(item.id)}>Approve</button>
          </div>
        ))}
      </div>
    );
  }
}
```

## 12. SPFx Version History

| Version | Release | Node.js | Key Features |
|---|---|---|---|
| 1.25 | Sept 2026 | 22 | GA: Copilot Components, Navigation Customizers, SPFx CLI |
| 1.24 | July-Aug 2026 | 22 | Preview: React 18, Copilot Components, SPFx CLI |
| 1.23 | May-June 2026 | 22 | ListView grouping, panel overrides |
| 1.22 | Dec 2025 | 22 | **Gulp → Heft toolchain**, TypeScript 5.8 |
| 1.21 | April 2025 | 18/22 | Node 22 support, flexible layouts |
| 1.19 | 2024 | 16-18 | ACE chart views |
| 1.16 | 2022-2023 | 16 | Search Query Modifiers |
| 1.15 | 2022 | 14-16 | Form Customizers |
| 1.13 | 2021 | 14 | ACEs / Viva Connections |
| 1.9.1 | 2019 | 10 | Library Components |
| 1.0 | 2017 | 6 | Initial GA: Web Parts + Extensions |

## 13. Performance Optimization

### Bundle Size Optimization

```bash
# Analyze bundle
gulp bundle --ship --analyze
```

**Use externals to prevent bundling large libraries:**

```json
// config/config.json
{
  "externals": {
    "react": "React",
    "react-dom": "ReactDOM",
    "@fluentui/react": {
      "path": "https://static2.sharepointonline.com/.../fluentui-react.min.js",
      "globalName": "FluentUIReact"
    }
  }
}
```

**Use lighter alternatives:**

```typescript
// WRONG: Full lodash (~70KB)
import _ from 'lodash';

// CORRECT: SPFx subset (~15KB)
import { cloneDeep } from '@microsoft/sp-lodash-subset';
```

### Code Splitting & Lazy Loading

```typescript
import React, { lazy, Suspense } from 'react';

const HeavyDataGrid = lazy(() => import('./components/DataGrid'));

const MyComponent = () => (
  <Suspense fallback={<Spinner label="Loading..." />}>
    <HeavyDataGrid />
  </Suspense>
);
```

### React Performance

```typescript
import React, { memo, useMemo, useCallback } from 'react';

// Memoize components
const TaskItem = memo(({ task, onSelect }) => (
  <div>
    <span>{task.title}</span>
    <button onClick={() => onSelect(task.id)}>Select</button>
  </div>
));

// Memoize expensive computations
const filteredTasks = useMemo(() => {
  return tasks.filter(t => t.status === filter).sort((a, b) => a.priority - b.priority);
}, [tasks, filter]);

// Stable callback references
const handleSelect = useCallback((id) => {
  console.log('Selected:', id);
}, []);
```

### Fluent UI Optimized Imports

```typescript
// WRONG: Imports entire package
import { Button, Spinner } from '@fluentui/react';

// CORRECT: Path imports
import { Button } from '@fluentui/react/lib/Button';
import { Spinner } from '@fluentui/react/lib/Spinner';
```

### Performance Metrics

| Metric | Target |
|---|---|
| LCP (Largest Contentful Paint) | < 2.5s |
| TBT (Total Blocking Time) | < 200ms |
| Bundle Size (main) | < 100KB gzipped |
| API Calls per render | Minimize |

## 14. Developer Tooling

### Environment Setup

```bash
# Install Node.js 22 LTS
nvm install 22
nvm use 22

# Verify
node --version
npm --version

# Install global tools (for Gulp-based projects)
npm install -g yo@5 generator-sharepoint@1.21.1 gulp-cli

# Trust dev certificate (first-time, as admin)
gulp trust-dev-cert

# Create new project
yo @microsoft/sharepoint

# Install and serve
npm install
gulp serve
```

### VS Code Extensions

| Extension | Purpose |
|---|---|
| SPFx Toolkit | Project management, scaffolding, deployment |
| Adaptive Card Previewer | Preview ACE templates |
| ESLint | Linting |
| Prettier | Code formatting |
| REST Client | Test SharePoint APIs |

### CLI for Microsoft 365

```bash
# Install
npm install -g @pnp/cli-microsoft365

# Authenticate
m365 login

# Upgrade project
m365 spfx project upgrade --toVersion 1.23.0

# Health check
m365 spfx project doctor

# Deploy
m365 spo app add --filePath ./solution/my-solution.sppkg --scope tenant
m365 spo app deploy --name my-solution-client-side-solution
```

### SharePoint Workbench

```
Local Dev Server: https://localhost:4321/
       +
SharePoint Workbench: https://[tenant].sharepoint.com/[site]/_layouts/workbench.aspx
       =
Your web part running against real SharePoint with real data
```

## 15. PnP Community & Resources

### PnP Community

| Resource | URL |
|---|---|
| SPFx Web Part Samples | github.com/pnp/sp-dev-fx-webparts |
| SPFx Extension Samples | github.com/pnp/sp-dev-fx-extensions |
| SPFx ACE Samples | github.com/pnp/sp-dev-fx-aces |
| PnPjs Library | pnp.github.io/pnpjs |
| PnP React Controls | pnp.github.io/sp-dev-fx-controls-react |
| PnP Property Controls | pnp.github.io/sp-dev-fx-property-controls |
| CLI for Microsoft 365 | pnp.github.io/cli-microsoft365 |

### Official Microsoft Resources

| Resource | URL |
|---|---|
| SPFx Documentation | learn.microsoft.com/sharepoint/dev/spfx |
| SPFx Release Notes | learn.microsoft.com/sharepoint/dev/spfx/release-notes |
| Microsoft 365 Dev Blog | devblogs.microsoft.com/microsoft365dev |
| Viva Connections Docs | learn.microsoft.com/viva/connections |
| Microsoft Graph Explorer | developer.microsoft.com/graph/graph-explorer |

## 16. Advanced Topics

### Unit Testing

```typescript
// src/__tests__/MyComponent.test.tsx
import * as React from 'react';
import { render, screen } from '@testing-library/react';
import { MyComponent } from '../MyComponent';

describe('MyComponent', () => {
  it('renders loading state initially', () => {
    const { container } = render(<MyComponent {...mockProps} />);
    expect(container.textContent).toContain('Loading');
  });

  it('renders items after data load', async () => {
    render(<MyComponent {...mockProps} />);
    const item = await screen.findByText('Test Task');
    expect(item).toBeTruthy();
  });
});
```

### CI/CD Pipeline

**GitHub Actions:**

```yaml
# .github/workflows/deploy-spfx.yml
name: SPFx CI/CD

on:
  push:
    branches: [main, develop]

jobs:
  build-and-test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      
      - name: Setup Node.js 22
        uses: actions/setup-node@v4
        with:
          node-version: '22'
          cache: 'npm'

      - run: npm ci
      - run: npm test -- --coverage --ci --watchAll=false
      - run: |
          gulp bundle --ship
          gulp package-solution --ship

      - uses: actions/upload-artifact@v4
        with:
          name: spfx-package
          path: sharepoint/solution/*.sppkg

  deploy-to-dev:
    needs: build-and-test
    if: github.ref == 'refs/heads/develop'
    runs-on: ubuntu-latest
    steps:
      - uses: actions/download-artifact@v4
      - run: |
          npm install -g @pnp/cli-microsoft365
          m365 login --authType certificate
          m365 spo app add --filePath ./*.sppkg --scope tenant --overwrite
          m365 spo app deploy --name my-solution --scope tenant
```

### Dynamic Data (Web Part Communication)

**Publisher Web Part:**

```typescript
import { IDynamicDataCallables } from '@microsoft/sp-component-base';

export default class FilterWebPart extends BaseClientSideWebPart<{}>
  implements IDynamicDataCallables {

  private _selectedFilter: string = '';

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve();
  }

  public getPropertyDefinitions() {
    return [
      { id: 'selectedFilter', title: 'Selected Filter Value' }
    ];
  }

  public getPropertyValue(propertyId: string): string {
    if (propertyId === 'selectedFilter') return this._selectedFilter;
    throw new Error(`Unknown property: ${propertyId}`);
  }

  private onFilterChange(newFilter: string): void {
    this._selectedFilter = newFilter;
    this.context.dynamicDataSourceManager.notifyPropertyChanged('selectedFilter');
  }
}
```

**Consumer Web Part:**

```typescript
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IDashboardProps {
  filterValue: DynamicProperty<string>;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardProps> {

  protected onInit(): Promise<void> {
    this.properties.filterValue.register(this.render.bind(this));
    return Promise.resolve();
  }

  public render(): void {
    const currentFilter = this.properties.filterValue.tryGetValue();
    ReactDom.render(
      React.createElement(Dashboard, { filterValue: currentFilter }),
      this.domElement
    );
  }
}
```

### Localization

```
src/webparts/myFeature/loc/
├── en-us.js
├── de-de.js
├── fr-fr.js
└── mystrings.d.ts
```

**en-us.js:**
```javascript
define([], function() {
  return {
    "LoadingMessage": "Loading data...",
    "SaveButton": "Save",
    "ErrorMessage": "An error occurred."
  };
});
```

**Usage:**
```typescript
import * as strings from 'MyFeatureWebPartStrings';

<button>{strings.SaveButton}</button>
```

### Accessibility (WCAG)

```typescript
// Semantic HTML
<button onClick={onDelete} aria-label={`Delete ${item.title}`}>
  <Icon iconName="Delete" aria-hidden="true" />
</button>

// Form validation
<input
  id="title-input"
  type="text"
  aria-required="true"
  aria-invalid={!!titleError}
  aria-describedby={titleError ? 'title-error' : undefined}
/>
{titleError && (
  <span id="title-error" role="alert">{titleError}</span>
)}

// Focus management in dialogs
useEffect(() => {
  if (isOpen && firstFocusableRef.current) {
    firstFocusableRef.current.focus();
  }
}, [isOpen]);
```

### Theming

```scss
// MyFeature.module.scss
.container {
  background: var(--neutralLighterAlt);
  color: var(--bodyText);
  border: 1px solid var(--neutralLight);
}

.header {
  color: var(--themePrimary);
  font-size: $ms-font-size-xl;
}
```

```typescript
// Respond to theme changes
protected onThemeChanged(currentTheme: IReadonlyTheme): void {
  const { semanticColors, palette } = currentTheme;
  this.domElement.style.setProperty('--themePrimary', palette.themePrimary);
}
```

### SharePoint Embedded

Use SharePoint as storage backend for external apps.

```typescript
// List containers
const client = await this.context.msGraphClientFactory.getClient('3');
const containers = await client
  .api('/storage/fileStorage/containers')
  .filter(`containerTypeId eq '${CONTAINER_TYPE_ID}'`)
  .get();

// Upload file
const containerDriveId = containers.value[0].id;
await client
  .api(`/drives/${containerDriveId}/root:/${file.name}:/content`)
  .put(file.arrayBuffer());
```

### Azure Functions Integration

```typescript
import { HttpClient } from '@microsoft/sp-http';

const options = {
  body: JSON.stringify(payload),
  headers: {
    'Content-Type': 'application/json',
    'x-functions-key': 'YOUR_KEY'
  }
};

const response = await this.context.httpClient.post(
  'https://my-function.azurewebsites.net/api/ProcessData',
  HttpClient.configurations.v1,
  options
);

const result = await response.json();
```

### Debugging

**VS Code Launch Config:**

```json
{
  "type": "msedge",
  "request": "launch",
  "url": "https://tenant.sharepoint.com/sites/site/_layouts/workbench.aspx",
  "webRoot": "${workspaceFolder}",
  "sourceMaps": true
}
```

**Bundle Analysis:**
```bash
gulp bundle --ship --analyze
```

**SPFx Doctor:**
```bash
m365 spfx project doctor
m365 spfx project upgrade --toVersion 1.23.0
```

### Common Issues

| Issue | Fix |
|---|---|
| Build fails | Check Node.js version matches SPFx requirements |
| `gulp serve` cert error | Run `gulp trust-dev-cert` as admin |
| 401 Graph API errors | Approve permissions in SP Admin → API access |
| Large bundle size | Use externals config; analyze bundle |
| Web part not loading | Check F12 console for JS errors |

---

**Official Resources:**
- Documentation: https://learn.microsoft.com/sharepoint/dev/spfx
- Samples: https://github.com/pnp/sp-dev-fx-webparts
- PnP Community: https://aka.ms/m365pnp
