# 📘 SPFx WebPart & Dashboard Development Guide
### For New Developers — Best Practices & Architecture Reference

> **Purpose:** A general-purpose guide for developing optimized, production-ready SharePoint Framework (SPFx) webparts and dashboards. Patterns and examples are grounded in the **Common Platform** reference project.
>
> **Reference Project:** `Common Platform` (SPFx 1.20.0)  
> **Stack:** SPFx 1.20.0 · React 17 · TypeScript 4.7.4 · PnP JS v4 · Node 18.x

---

## 🚀 Executive Development Summary (Developer Cheat Sheet & Priorities)

Before diving into full features or writing code for a new SPFx project, every developer **MUST** understand and prioritize these core rules. This section ranks development priorities by **technical weight, architecture impact, and failure risk**.

```
    PRIORITY 1: Core Toolchain & Environment (Node / SPFx / Heft vs Gulp)
                              │
    PRIORITY 2: Solution Architecture & Reusable Framework Layer
                              │
    PRIORITY 3: Data Access & State Lifecycle (PnP Singleton & React Mount)
                              │
    PRIORITY 4: Production Resilience (Logging, Validation & Batching)
                              │
    PRIORITY 5: Graph API & Platform Security Permissions
```

---

### Priority 1: Toolchain & Build Environment (Highest Technical Heft)
- **Toolchain Alignment:** Match Node.js & SPFx versions strictly.
  - **SPFx v1.20.0:** Requires **Node v18.x** + **Gulp** (`gulp bundle --ship`).
  - **SPFx v1.22.0+:** Requires **Node v18/v20.x** + **Heft** (`npx heft build --production`).
- **Dev Certificate:** Always run `gulp trust-dev-cert` (or `npx heft trust-dev-cert`) before debugging locally.
- **Node Mismatch Risk:** Node version mismatch is the **#1 root cause of build failures**. Use `nvm use 18` prior to scaffolding or running tasks.

---

### Priority 2: Shared Architecture Layer (`src/framework/`)
- **Strict Separation of Concerns:**
  - `src/framework/`: Shared across all webparts (Services, Utilities, Constants).
  - `src/webparts/<name>/components/`: WebPart-specific UI only.
- **Centralized Constants:** Never hardcode list names, library names, or site relative paths in component files. Store all string identifiers in `src/framework/Constants/Constant.ts`.
- **Golden Rule:** If code can be reused by another WebPart or feature, place it inside `src/framework/`.

---

### Priority 3: Data Operations & React Lifecycle Rules
- **PnP Singleton (`pnpjsConfig.ts`):** Instantiate `SPFI` **exactly once** using `getSP(context)`. Never instantiate PnP inside constructors or component render methods.
- **Component Lifecycle:** 
  - `constructor()`: Context is `undefined`. Do **NOT** call APIs or initialize SP services here.
  - `componentDidMount()`: Initialize PnP/Utilities services and fetch initial state here.
  - `render()`: Keep standard null guards (`this._utilities ? ... : fallback`) because `render()` fires before `componentDidMount()` finishes async calls.
  - `onDispose()`: **Always** execute `ReactDom.unmountComponentAtNode(this.domElement)` to prevent memory leaks when webparts unmount.
- **Server-Side Filtering & Selection:**
  - Always chain `.select(...)` to fetch only required fields (never fetch all 40+ default columns).
  - Use server-side `.filter(...)` (OData) instead of JavaScript `.filter()` post-fetch.
  - For lists with $>5,000$ items, use the async iterator pattern (`GetAllItemsFromList`).

---

### Priority 4: Data Mutation, Validation & Logging Standard
- **Batching Writes:** When creating/updating multiple items, use batch requests (`addMultipleItemsToList`, `updateListItemsInBatch`) to consolidate HTTP roundtrips.
- **Recycle Bin over Hard Delete:** Use `recycle()` instead of hard `delete()` to protect customer data.
- **People Picker Fields:** Always save SharePoint People Picker values using resolved integer User IDs (`AssignedToId`), **never** raw email strings.
- **Production Logging:** All caught service errors **MUST** log to the SharePoint `LogHistory` list using `addItemsToLogList()` (with exact `functionName` and `pageName`).

---

### Priority 5: Graph API Security & Deployment
- **Permissions Declaration:** Declare all Microsoft Graph scopes in `config/package-solution.json` (`webApiPermissionRequests`).
- **Admin Approval:** After deploying `.sppkg`, approve pending permissions in **SharePoint Admin Center $\rightarrow$ Advanced $\rightarrow$ API access**.
- **Graph Client Version:** Always call `context.msGraphClientFactory.getClient("3")` (use version `"3"`).
- **Pagination:** Graph returns max 100-999 items per request; always loop through `@odata.nextLink`.

---

## 📑 Table of Contents

1. [Executive Development Summary (Developer Cheat Sheet & Priorities)](#-executive-development-summary-developer-cheat-sheet--priorities)
2. [What is SPFx? Key Concepts for New Developers](#1-what-is-spfx-key-concepts-for-new-developers)
3. [Prerequisites & Environment Setup](#2-prerequisites--environment-setup)
4. [SPFx Project Setup & Scaffolding](#3-spfx-project-setup--scaffolding)
5. [Recommended Project Architecture](#4-recommended-project-architecture)
6. [Folder Structure & Naming Conventions](#5-folder-structure--naming-conventions)
7. [PnP JS — The Right Way to Talk to SharePoint](#6-pnp-js--the-right-way-to-talk-to-sharepoint)
8. [SharePoint CRUD Operations — Patterns & Best Practices](#7-sharepoint-crud-operations--patterns--best-practices)
9. [Graph API — Connecting to Microsoft 365](#8-graph-api--connecting-to-microsoft-365)
10. [WebPart Lifecycle — Understand Before You Code](#9-webpart-lifecycle--understand-before-you-code)
11. [React Component Architecture in SPFx](#10-react-component-architecture-in-spfx)
12. [Building Forms — State, Validation & Handlers](#11-building-forms--state-validation--handlers)
13. [Building Dashboards — Tables, Filters & Exports](#12-building-dashboards--tables-filters--exports)
14. [Error Handling & Logging — Never Skip This](#13-error-handling--logging--never-skip-this)
15. [Performance Optimization Rules](#14-performance-optimization-rules)
16. [Common SPFx Pitfalls to Avoid](#15-common-spfx-pitfalls-to-avoid)
17. [Build, Bundle & Deploy Workflow](#16-build-bundle--deploy-workflow)
18. [Code Quality Checklist](#17-code-quality-checklist)
19. [Quick Reference](#18-quick-reference)

---

## 1. What is SPFx? Key Concepts for New Developers

**SharePoint Framework (SPFx)** is Microsoft's recommended development model for building client-side experiences in SharePoint Online, Microsoft Teams, and Viva Connections. Unlike classic SharePoint development, SPFx runs entirely in the browser — there is no server-side code execution.

### How SPFx Works — The Mental Model

```
SharePoint Page
    │
    ├── Page hosts one or more WebParts (your app)
    │       │
    │       ├── WebPart renders a React component tree
    │       │       │
    │       │       ├── Components call services (PnP, Graph API)
    │       │       └── Services communicate with SharePoint / Microsoft 365
    │       │
    │       └── WebPart receives context (site URL, user info, httpClient, etc.)
    │
    └── Property Pane (right-side panel) — configuration UI for the WebPart
```

### Key SPFx Concepts You Must Know

| Concept | What It Means |
|---|---|
| **WebPart** | A self-contained React application deployed to SharePoint |
| **WebPartContext** | The object SharePoint gives you — contains site info, current user, API clients |
| **BaseClientSideWebPart** | The base class every webpart extends |
| **Property Pane** | Right-side configuration panel for end-users |
| **manifest.json** | Defines the webpart ID, title, icon, and supported hosts |
| **package-solution.json** | Defines solution metadata, Graph API permissions, and deployment options |
| **SPFI (PnP)** | The PnP JS client instance used for all SharePoint REST calls |
| **MSGraphClient** | The client used for Microsoft Graph API calls |
| **SCSS Modules** | CSS files scoped to the component to prevent style leakage |

### Important: SPFx Has Rules

- No server-side code. All logic runs in the browser.
- You cannot use Node.js modules that depend on `fs`, `path`, etc.
- API calls go through SharePoint's `httpClient` or PnP JS — never raw `axios`/`fetch` for SP lists.
- All external NPM packages must be browser-compatible.
- The `context` object (containing user info, site URLs, API clients) is only available **after** the webpart renders — not in the constructor.

---

## 2. Prerequisites & Environment Setup

Setting up a robust development environment is critical for SPFx stability. Follow the official Microsoft Learn guidelines for configuring your developer machine.

### Required Software & Toolchain Matrix

SPFx versions are tightly coupled with specific Node.js and package manager versions. Reference the compatibility matrix below:

| SPFx Version | Node.js Version | Recommended NPM | Primary Build Toolchain |
|---|---|---|---|
| **v1.20.0 (Current)** | `v18.x` (LTS) | `v8.x` or `v9.x` | Gulp / gulp-cli |
| **v1.22.0+ (Modern)** | `v18.x` / `v20.x` | `v10.x` | **Heft** (Rush Stack task orchestrator) |

> [!WARNING]
> Node.js version mismatches are the #1 cause of build errors. Verify your local configuration using `node -v` and `npm -v` before running setup.

### Development Environment Setup Step-by-Step

Follow these steps to configure your global environment:

1. **Install Node.js (Version Manager Recommended)**
   Use `nvm` (Node Version Manager) for Windows or macOS to install and switch to Node.js v18.x:
   ```bash
   nvm install 18
   nvm use 18
   ```

2. **Install Global Build Tooling**
   For projects using legacy/current SPFx builds, install Gulp and Yeoman globally:
   ```bash
   npm install -g gulp-cli yo
   ```

3. **Install the Yeoman SharePoint Generator**
   Ensure you have the generator that matches your project version:
   ```bash
   npm install -g @microsoft/generator-sharepoint@1.20.0
   ```

4. **Trust the Self-Signed Developer Certificate (Crucial for Local Debugging)**
   To preview webparts in a live SharePoint page or the local workbench, your browser must trust the SPFx developer SSL certificate. Run the following command:
   ```bash
   gulp trust-dev-cert
   ```
   *For modern Heft-based projects (v1.22+), run:*
   ```bash
   npx heft trust-dev-cert
   ```
   Select **Yes** when prompted to install the certificate.

### Modern vs. Legacy Toolchain: Heft vs. Gulp

Starting in SPFx v1.22, Microsoft transitioned from the Gulp-based build toolchain to **Heft** (from the Rush Stack ecosystem).

* **Gulp Toolchain (Legacy / v1.20.0):** Relies on a global `gulp-cli` installation and a local `gulpfile.js` to execute tasks like bundling, packaging, and cleaning.
* **Heft Toolchain (Modern / v1.22.0+):** Acts as a pluggable build task orchestrator. It uses config-driven phases (`heft.json`) and rig systems (`@microsoft/spfx-rig`) instead of custom programmatic build files, significantly reducing technical debt and improving performance through parallel task execution.

---

## 3. SPFx Project Setup & Scaffolding

### Creating a New SPFx Project

```bash
# Navigate to your project folder
cd C:\Projects\MyProject

# Run the generator
yo @microsoft/sharepoint
```

**Generator Prompts — What to Select:**

| Prompt | Recommended Choice | Notes |
|---|---|---|
| Solution name | `my-project` | Use kebab-case |
| Target environment | `SharePoint Online only` | We deploy to SP Online |
| Place files in subfolder? | `No` | Keep flat |
| Tenant admin deployment? | `Yes` | Allows global deploy |
| Type of component | `WebPart` | Start with a webpart |
| WebPart name | `MyFeature` | PascalCase |
| Template | `React` | Always use React |

### Adding a New WebPart to an Existing Project

```bash
# Inside the existing project folder
yo @microsoft/sharepoint

# Select: Add a new web part
# Choose React as the framework
```

This adds the new webpart to the existing solution — both webparts are bundled and deployed together as one `.sppkg` file.

### Install Core Dependencies

After scaffolding, install the essential packages used in the Common Platform reference:

```bash
# PnP JS (SharePoint operations)
npm install @pnp/sp @pnp/logging @pnp/graph @pnp/queryable

# Fluent UI (Microsoft's design system for SharePoint)
npm install @fluentui/react

# Date handling
npm install moment

# Table component for dashboards
npm install material-react-table

# Export utilities
npm install export-to-csv xlsx

# PDF generation
npm install jspdf html2canvas

# Toast notifications
npm install react-toastify
```

---

## 4. Recommended Project Architecture

### The Layered Architecture Pattern

The most maintainable SPFx projects separate concerns into distinct layers. This is the pattern used in the Common Platform reference and is the recommended standard for enterprise SPFx projects:

```
src/
├── framework/             ← SHARED LAYER (reusable across all webparts)
│   ├── Constants/         ← All hardcoded values (list names, etc.)
│   ├── Services/          ← Business logic, API calls
│   └── Utilities/         ← Formatting, export, helper functions
│
├── pnpjsConfig.ts         ← PnP SP instance (singleton — one per solution)
│
└── webparts/              ← WEBPART LAYER (specific UI & wiring)
    └── myWebPart/
        ├── MyWebPart.ts               ← Entry point only
        └── components/
            ├── IMyWebPartProps.ts     ← Props contract
            ├── MyWebPart.tsx          ← Root component (thin)
            ├── MyFeature.tsx          ← Business logic + UI
            └── MyWebPart.module.scss  ← Scoped styles
```

### Why This Architecture?

**Without a shared framework layer:**
- Every webpart duplicates SP CRUD code
- A list rename breaks N files
- Testing requires deploying each webpart separately
- New developers write inconsistent code

**With a shared framework layer:**
- One place to fix a bug — all webparts benefit
- List names in one file — one change propagates everywhere
- New developers follow established patterns
- Consistent error handling and logging across all webparts

### The Golden Rule

> **"If two webparts would need the same code, it belongs in `framework/`. If it is specific to one webpart's UI or behavior, it belongs in that webpart's `components/` folder."**

---

## 5. Folder Structure & Naming Conventions

### Full Reference Structure

```
YourProject/
│
├── config/
│   ├── config.json               # Bundle entries — auto-updated by generator
│   ├── package-solution.json     # ⭐ IMPORTANT: Graph permissions + solution ID/version
│   ├── serve.json                # Local dev server URL configuration
│   ├── write-manifests.json      # CDN path (for Azure blob deploy)
│   └── sass.json                 # SCSS configuration
│
├── src/
│   ├── index.ts                  # Required by TS compiler — leave empty
│   ├── pnpjsConfig.ts            # ⭐ PnP singleton setup — one file for the whole solution
│   │
│   ├── framework/
│   │   ├── Constants/
│   │   │   └── Constant.ts       # ALL list names, library names, folder names
│   │   ├── Services/
│   │   │   └── ListOperations/
│   │   │       ├── PnpListOperationService.ts   # ALL SharePoint CRUD
│   │   │       ├── GraphAPIOperationService.ts  # ALL Microsoft Graph calls
│   │   │       ├── ValidationService.ts         # ALL form field validation
│   │   │       └── FormHandlers.ts              # Generic form state helpers
│   │   └── Utilities/
│   │       ├── CommonUtilities.ts   # Date, number, string, email, PDF, API
│   │       ├── ExcelUtilities.ts    # Excel file reading
│   │       └── CSVExportService.ts  # CSV file download
│   │
│   └── webparts/
│       └── myFeature/
│           ├── MyFeatureWebPart.ts                 # Entry — extends BaseClientSideWebPart
│           ├── MyFeatureWebPart.manifest.json       # WebPart identity & metadata
│           ├── assets/                             # Images, icons
│           ├── loc/                                # Localization string files
│           └── components/
│               ├── IMyFeatureProps.ts              # Props interface
│               ├── MyFeature.tsx                   # Root component (thin wrapper)
│               ├── MyFeatureMain.tsx               # Business logic component
│               └── MyFeature.module.scss           # Scoped SCSS
│
├── package.json
├── tsconfig.json
├── gulpfile.js
└── .eslintrc.js
```

### Naming Conventions — Follow These Always

| Item | Convention | Correct Example | Wrong Example |
|---|---|---|---|
| WebPart folder | `camelCase` | `myFeature` | `MyFeature`, `my-feature` |
| WebPart entry file | `PascalCase + WebPart.ts` | `MyFeatureWebPart.ts` | `myFeature.ts` |
| Root component | `PascalCase.tsx` | `MyFeature.tsx` | `myfeature.tsx` |
| Business component | `PascalCase.tsx` | `MyFeatureMain.tsx` | `main.tsx` |
| Props interface file | `I + PascalCase + Props.ts` | `IMyFeatureProps.ts` | `props.ts` |
| Style file | `PascalCase + .module.scss` | `MyFeature.module.scss` | `styles.scss` |
| Service classes | `I + PascalCase + Service` | `IPnpListOperationService` | `ListService` |
| Utility classes | `I + PascalCase + Utilities` | `ICommonUtilities` | `utils.ts` |
| Constants class | `PascalCase` | `Constant` | `constants.ts` |
| Interface names | `I + PascalCase` | `IFormState` | `FormState` |

---

## 6. PnP JS — The Right Way to Talk to SharePoint

### What is PnP JS?

PnP (Patterns and Practices) JS is a **wrapper library** around SharePoint's REST APIs. Instead of writing raw `fetch()` calls with complex URL strings, PnP lets you write clean, chainable TypeScript.

```typescript
// Without PnP — raw REST (verbose, error-prone)
const response = await fetch(
  `${siteUrl}/_api/web/lists/getbytitle('MyList')/items?$select=Id,Title&$filter=Status eq 'Active'`,
  { headers: { "Accept": "application/json;odata=verbose" } }
);

// With PnP — clean and readable
const items = await sp.web.lists.getByTitle('MyList').items
  .filter("Status eq 'Active'")
  .select('Id', 'Title')();
```

### The Singleton Pattern — One Instance for All

Never create multiple PnP instances. One instance per solution, initialized once with context, then reused everywhere.

**`src/pnpjsConfig.ts`** (reference implementation):
```typescript
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";

// Module-level singleton — shared across all services in this solution
let _sp: SPFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context !== null) {
    // Create the instance exactly ONCE — on first call with context
    _sp = spfi()
      .using(SPFx(context))                    // Wires SPFx auth + site URL
      .using(PnPLogging(LogLevel.Warning));    // Log warnings/errors to console
  }
  return _sp;
};
```

### Why Singleton?

- Prevents multiple auth handshakes
- Consistent behavior across all service calls
- Single point to change configuration (log level, retries, etc.)
- Resets automatically on page reload — no manual cleanup needed

### Required PnP Imports (add to your service file as needed)

```typescript
// Core
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

// For file/folder operations
import "@pnp/sp/files";
import "@pnp/sp/folders";

// For field metadata
import "@pnp/sp/fields";

// For site users/groups
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";

// For sending email via SharePoint utility
import "@pnp/sp/sputilities";

// For batching multiple operations
import "@pnp/sp/batching";

// For regional settings (timezone)
import "@pnp/sp/regional-settings/web";
```

> **Rule:** Only import what you use. Each import adds to your bundle size.

---

## 7. SharePoint CRUD Operations — Patterns & Best Practices

### Service Initialization Pattern

Every service class that talks to SharePoint must:
1. Store the `SPFI` instance as a class field
2. Expose an `Init(context)` method that calls `getSP(context)`
3. Never accept or store `context` directly — only the `SPFI` instance

```typescript
// Reference: PnpListOperationService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../../../pnpjsConfig";

export class IPnpListOperationService {
    public _sp: SPFI = null;

    // Always call this before using any other method
    public async Init(context?: WebPartContext): Promise<void> {
        this._sp = getSP(context);
    }
}
```

---

### READ — Fetching Items from a List

#### Fetching All Items (Handles Lists > 5000 Items)

Use the **async iterator** pattern with `GetAllItemsFromList`. This correctly handles SharePoint's list view threshold (5000 item limit per request):

```typescript
// Correct pattern — uses async iterator, handles pagination
public async GetAllItemsFromList(
    listName: string,
    filter: string = "",
    select: string[] = [],
    expand: string[] = []
): Promise<any[]> {
    const allItems: any[] = [];
    try {
        let items = this._sp.web.lists.getByTitle(listName).items;

        if (filter) items = items.filter(filter);
        if (select.length > 0) items = items.select(...select);
        if (expand.length > 0) items = items.expand(...expand);

        // Async iterator — fetches page by page until all items are retrieved
        for await (const page of items.top(5000)) {
            allItems.push(...page);
        }

        return allItems;
    } catch (error) {
        // Always log errors (see Section 13)
        throw error;
    }
}
```

**Usage:**
```typescript
// Always use select() to limit fields — never fetch all columns blindly
const items = await _listService.GetAllItemsFromList(
    Constant.listName,
    "Status eq 'Active'",             // OData filter
    ["Id", "Title", "Status", "Author/Title"],  // Only the fields you need
    ["Author"]                         // Expand lookup fields
);
```

#### Fetching With Ordering

```typescript
const items = await _listService.getItemsFromList(
    Constant.listName,
    "",                       // No filter
    ["Id", "Title", "Created"],
    [],                       // No expand
    "Created",                // Order by field
    false                     // false = descending (newest first)
);
```

#### Fetching a Single Item by ID

```typescript
const item = await _listService.getItemById(
    Constant.listName,
    itemId,
    ["Id", "Title", "Status", "AssignedTo/Title"],
    ["AssignedTo"]
);
```

---

### WRITE — Adding Items

#### Single Item

```typescript
// The item object keys must match the SharePoint internal column names
const newItem = {
    Title: formState.title,
    Status: "Draft",
    Amount: formState.amount,
    StartDate: formState.startDate,
    // People Picker: must be integer ID (not email)
    AssignedToId: userId
};

await _listService.addItemsToList(Constant.listName, newItem);
```

#### Multiple Items — Use Batching

```typescript
// Efficient: one HTTP request for all items
const items = [
    { Title: "Item A", Status: "Draft" },
    { Title: "Item B", Status: "Draft" },
    { Title: "Item C", Status: "Draft" }
];

// Sends all 3 in a single batched HTTP request
await _listService.addMultipleItemsToList(Constant.listName, items);

// If you need the IDs of created items back:
const createdIds = await _listService.addMultipleItemsToListForIteration(
    Constant.listName, items
);
// createdIds = [101, 102, 103]
```

---

### UPDATE — Modifying Items

#### Single Item

```typescript
// Only include fields that are changing
await _listService.updateItemInList(Constant.listName, itemId, {
    Status: "Approved",
    ApprovedDate: new Date().toISOString()
});
```

#### Multiple Items — Use Batching

```typescript
const updates = [
    { id: 101, fields: { Status: "Approved" } },
    { id: 102, fields: { Status: "Rejected" } },
    { id: 103, fields: { Status: "Pending" } }
];

// All 3 updates in one batched request
await _listService.updateListItemsInBatch(Constant.listName, updates);
```

---

### DELETE — Removing Items

```typescript
// ALWAYS use recycle() — sends to Recycle Bin (recoverable)
// Never use delete() — this is a permanent hard delete
await _listService.deleteItemFromList(Constant.listName, itemId);
```

---

### FILE & FOLDER OPERATIONS

#### Upload a File

```typescript
// Upload to library root
const serverRelativeUrl = await _listService.uploadFileToLibraryRoot(
    file,                           // File object (from file input)
    "/sites/MySite/Shared Documents" // Full server-relative path
);

// Upload to a specific folder
await _listService.uploadFileToLibraryFolder(
    file,
    "/sites/MySite/Shared Documents",
    "2026/Q1"                        // Folder path within library
);

// Upload + update metadata in one step
await _listService.uploadFileToLibraryFolderWithUpdateMetaData(
    file,
    "/sites/MySite/Shared Documents",
    "2026/Q1",
    { Title: "My Document", Department: "Finance" }  // Metadata
);
```

#### Working With Folders

```typescript
// Check if folder exists before creating
const exists = await _listService.CheckFolderExist(
    "/sites/MySite/Shared Documents/2026"
);
if (!exists) {
    await _listService.createFolder("Shared Documents", "2026");
}

// Get all files in a folder
const files = await _listService.getFilesFromFolder(
    "/sites/MySite/Shared Documents/2026/Q1"
);

// Get latest file
const latestFile = await _listService.getLatestFileFromFolder(
    "/sites/MySite/Shared Documents/2026/Q1"
);
```

---

### USER & GROUP OPERATIONS

```typescript
// Get current user's ID
const currentUser = await _listService.getCurrentUser();
const userId = currentUser.Id;

// Resolve user ID from email (needed for People Picker fields)
const userId = await _listService.getUserId("john.doe@company.com");

// Get all SP groups
const groups = await _listService.getAllSiteGroups();

// Check if current user is in a specific SP group
const isAdmin = await _listService.isCurrentUserInGroup("Site Administrators");

// Get all users in a specific SP group
const approvers = await _listService.getUsersFromGroupByName("Approvers");
```

---

## 8. Graph API — Connecting to Microsoft 365

### When to Use Graph API vs PnP

| Use Case | Use |
|---|---|
| Read/write SharePoint list items | PnP JS |
| Upload files to document library | PnP JS |
| Get Azure AD users (not SP users) | Graph API |
| Get Azure AD groups | Graph API |
| Send email via Graph | Graph API |
| Access Teams channels | Graph API |
| Get OneDrive/SharePoint via Graph | Graph API |
| Audit logs, risk events | Graph API |

### Step 1: Declare Permissions

Before writing a single line of Graph code, add the required permissions to `config/package-solution.json`:

```json
{
  "solution": {
    "webApiPermissionRequests": [
      { "resource": "Microsoft Graph", "scope": "User.Read.All" },
      { "resource": "Microsoft Graph", "scope": "GroupMember.Read.All" },
      { "resource": "Microsoft Graph", "scope": "Mail.Send" }
    ]
  }
}
```

> ⚠️ If you add permissions without this declaration, Graph calls will fail with `401 Unauthorized` in production.

### Step 2: Approve Permissions After Deployment

After deploying the `.sppkg` to the App Catalog:
1. Go to **SharePoint Admin Center → Advanced → API access**
2. Find your pending requests
3. Click **Approve**

### Step 3: Make Graph Calls

The Graph service pattern passes `context` per-method (unlike PnP which stores the SPFI instance):

```typescript
export class IGraphAPIOperationService {

    // Always get a fresh client — version "3" is required for modern Graph endpoints
    public async getAllUsers(context: WebPartContext): Promise<any[]> {
        const client = await context.msGraphClientFactory.getClient("3");
        let users: any[] = [];
        let nextLink: string = "/users?$top=999&$orderby=displayName asc";

        try {
            // Graph returns paginated results — ALWAYS handle @odata.nextLink
            while (nextLink) {
                const response = await client.api(nextLink).get();
                users = [...users, ...response.value];
                // If more pages exist, nextLink has the URL; otherwise it's null
                nextLink = response["@odata.nextLink"] || null;
            }
            return users;
        } catch (error) {
            console.error("Error fetching users:", error);
            return [];
        }
    }
}
```

**Usage in a component:**
```typescript
private _graphService: IGraphAPIOperationService = null;

async componentDidMount(): Promise<void> {
    // Graph service: no Init() needed — context passed per call
    this._graphService = new IGraphAPIOperationService();
    const users = await this._graphService.getAllUsers(this.props.context);
}
```

### Graph API Permission Reference

```
Get all AD users             → User.Read.All
Get current user             → User.Read
Get disabled accounts        → User.Read.All
Get user's groups            → GroupMember.Read.All, Directory.Read.All
Get all AD groups            → GroupMember.Read.All
Get group members            → GroupMember.Read.All
Send email                   → Mail.Send
Send email on behalf         → Mail.Send + Send.Mail.All
Access SP sites via Graph    → Sites.Read.All
Read/write SP via Graph      → Sites.ReadWrite.All
Audit logs                   → AuditLog.Read.All
Risk events                  → IdentityRiskEvent.Read.All
License info                 → Directory.Read.All
Teams channels               → Channel.ReadBasic.All
Teams channel members        → ChannelMember.Read.All
```

> **Principle of Least Privilege:** Only request the permissions your solution actually needs. Do not add permissions "just in case."

---

## 9. WebPart Lifecycle — Understand Before You Code

Every SPFx webpart has a defined lifecycle. Understanding it prevents the most common bugs new developers face.

### Lifecycle Order

```
1. Constructor (sync)        ← context NOT available here
       ↓
2. onInit() (async)          ← first async hook, context available
       ↓
3. render() (sync)           ← renders React component into DOM
       ↓
4. [React] componentDidMount ← BEST PLACE to initialize services & fetch data
       ↓
5. onThemeChanged()          ← whenever SP theme changes
       ↓
6. render() again            ← on property pane changes
       ↓
7. onDispose()               ← webpart removed from page — ALWAYS unmount React
```

### The Complete WebPart Entry File

```typescript
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import MyFeature from './components/MyFeature';
import { IMyFeatureProps } from './components/IMyFeatureProps';

// Define all configurable properties (available in Property Pane)
export interface IMyFeatureWebPartProps {
  description: string;
  listName: string;
  showExportButton: boolean;
  newFormUrl: string;
}

export default class MyFeatureWebPart
  extends BaseClientSideWebPart<IMyFeatureWebPartProps> {

  // ── RENDER ──────────────────────────────────────────────────────────────
  // Called every time the webpart needs to render (including property pane changes)
  // IMPORTANT: Do NOT put business logic here. Only create the React element.
  public render(): void {
    const element: React.ReactElement<IMyFeatureProps> = React.createElement(
      MyFeature,
      {
        context: this.context,                          // Required by PnP + Graph services
        listName: this.properties.listName,
        showExportButton: this.properties.showExportButton,
        newFormUrl: this.properties.newFormUrl,
        myhttpclient: this.context.httpClient          // Required for external API calls
      }
    );
    ReactDom.render(element, this.domElement);
  }

  // ── ON INIT ─────────────────────────────────────────────────────────────
  // Called once during webpart initialization. Return a Promise<void>.
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(_message => {
      // Optional: detect environment (SharePoint / Teams / Office)
    });
  }

  // ── ON THEME CHANGED ────────────────────────────────────────────────────
  // Fires when the SharePoint site theme changes.
  // Apply theme colors as CSS custom properties.
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  // ── ON DISPOSE ──────────────────────────────────────────────────────────
  // Called when the webpart is removed from the page.
  // ALWAYS unmount React to prevent memory leaks.
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // ── DATA VERSION ────────────────────────────────────────────────────────
  // Increment this when you change the shape of IMyFeatureWebPartProps.
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // ── PROPERTY PANE ───────────────────────────────────────────────────────
  // The right-side configuration panel shown to page editors.
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "WebPart Configuration" },
          groups: [
            {
              groupName: "Data Settings",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'List Name',
                  placeholder: 'Enter SharePoint list name'
                }),
                PropertyPaneTextField('newFormUrl', {
                  label: 'New Form URL'
                })
              ]
            },
            {
              groupName: "Display Settings",
              groupFields: [
                PropertyPaneToggle('showExportButton', {
                  label: 'Show Export Button',
                  onText: 'Visible',
                  offText: 'Hidden'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // ── PRIVATE HELPER ──────────────────────────────────────────────────────
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          switch (context.app.host.name) {
            case 'Office': return this.context.isServedFromLocalhost ? 'Local Office' : 'Office';
            case 'Outlook': return this.context.isServedFromLocalhost ? 'Local Outlook' : 'Outlook';
            case 'Teams':
            case 'TeamsModern': return this.context.isServedFromLocalhost ? 'Local Teams' : 'Teams';
            default: return 'Unknown environment';
          }
        });
    }
    return Promise.resolve(
      this.context.isServedFromLocalhost ? 'Local SharePoint' : 'SharePoint Online'
    );
  }
}
```

---

## 10. React Component Architecture in SPFx

### The 3-Layer Pattern

Every webpart's React tree should follow this pattern:

```
Layer 1: WebPart.ts         (Entry — wires SP context to React)
    ↓
Layer 2: RootComponent.tsx  (Thin wrapper — passes props, initializes shared state)
    ↓
Layer 3: MainComponent.tsx  (Business logic + UI — where real work happens)
```

### Layer 1 — WebPart Entry (already covered in Section 9)

### Layer 2 — Root Component (Thin Wrapper)

```typescript
// MyFeature.tsx — Root component
// Purpose: Thin wrapper. Pass props down. Optional: fetch shared state.
import * as React from 'react';
import type { IMyFeatureProps } from './IMyFeatureProps';
import MyFeatureMain from './MyFeatureMain';
import { ICommonUtilities } from '../../../framework/Utilities/CommonUtilities';

export default class MyFeature extends React.Component<IMyFeatureProps> {
  private _utilities: ICommonUtilities = null;
  private _timezoneOffset: number = 0;

  async componentDidMount(): Promise<void> {
    // Initialize shared utilities if needed at the root level
    this._utilities = new ICommonUtilities();
    await this._utilities.Init(this.props.context);
    // Fetch shared data (e.g., timezone) once and pass to children
    this._timezoneOffset = await this._utilities.getSPCurrentTimeZone();
    this.forceUpdate(); // Trigger re-render after async init
  }

  public render(): React.ReactElement<IMyFeatureProps> {
    return (
      <div>
        <MyFeatureMain
          context={this.props.context}
          listName={this.props.listName}
          timezoneOffset={this._timezoneOffset}
          showExportButton={this.props.showExportButton}
        />
      </div>
    );
  }
}
```

### Layer 3 — Business Component (Full Pattern)

```typescript
// MyFeatureMain.tsx — Where all the real work happens
import * as React from 'react';
import { IPnpListOperationService } from '../../../framework/Services/ListOperations/PnpListOperationService';
import { ValidationService } from '../../../framework/Services/ListOperations/ValidationService';
import { ICommonUtilities } from '../../../framework/Utilities/CommonUtilities';
import { CSVExportService } from '../../../framework/Utilities/CSVExportService';
import { Constant } from '../../../framework/Constants/Constant';

// Define State interface — always type your state
interface IMyFeatureState {
  items: IMyListItem[];
  isLoading: boolean;
  errorMessage: string;
  formState: IFormState;
  validationErrors: IValidationErrors;
}

// Define data models — avoid 'any' for list items
interface IMyListItem {
  Id: number;
  Title: string;
  Status: string;
  Amount: number;
}

export class MyFeatureMain extends React.Component<IMyFeatureMainProps, IMyFeatureState> {
  // Declare service instances as null — initialized in componentDidMount
  private _listService: IPnpListOperationService = null;
  private _validator: ValidationService = null;
  private _utilities: ICommonUtilities = null;

  constructor(props: IMyFeatureMainProps) {
    super(props);
    this.state = {
      items: [],
      isLoading: true,
      errorMessage: '',
      formState: { title: '', status: 'Draft', amount: 0 },
      validationErrors: {}
    };
  }

  // ── LIFECYCLE ────────────────────────────────────────────────────────
  async componentDidMount(): Promise<void> {
    // 1. Initialize services
    this._listService = new IPnpListOperationService();
    await this._listService.Init(this.props.context);

    this._validator = new ValidationService();
    // Note: ValidationService and GraphAPIOperationService do NOT need Init()

    this._utilities = new ICommonUtilities();
    await this._utilities.Init(this.props.context);

    // 2. Load initial data
    await this.loadItems();
  }

  // ── DATA OPERATIONS ───────────────────────────────────────────────────
  private async loadItems(): Promise<void> {
    try {
      this.setState({ isLoading: true, errorMessage: '' });
      const items = await this._listService.GetAllItemsFromList(
        Constant.listName,
        "Status ne 'Deleted'",
        ["Id", "Title", "Status", "Amount"],
        []
      );
      this.setState({ items, isLoading: false });
    } catch (error) {
      this.setState({ isLoading: false, errorMessage: 'Failed to load data.' });
    }
  }

  private async onSave(): Promise<void> {
    if (!this.validateForm()) return;
    try {
      await this._listService.addItemsToList(Constant.listName, {
        Title: this.state.formState.title,
        Status: this.state.formState.status,
        Amount: this.state.formState.amount
      });
      await this.loadItems(); // Refresh
    } catch (error) {
      this.setState({ errorMessage: 'Failed to save item.' });
    }
  }

  // ── VALIDATION ────────────────────────────────────────────────────────
  private validateForm(): boolean {
    let isValid = true;
    const errors: any = {};

    const [titleErr, titleMsg] = this._validator.isTextFieldEmpty(
      this.state.formState.title, ""
    );
    if (titleErr) { errors.title = titleMsg; isValid = false; }

    const [amountErr, amountMsg] = this._validator.isNumberFieldEmpty(
      this.state.formState.amount, ""
    );
    if (amountErr) { errors.amount = amountMsg; isValid = false; }

    this.setState({ validationErrors: errors });
    return isValid;
  }

  // ── RENDER ────────────────────────────────────────────────────────────
  public render(): React.ReactElement {
    const { isLoading, items, errorMessage } = this.state;

    if (isLoading) return <div>Loading...</div>;
    if (errorMessage) return <div style={{ color: 'red' }}>{errorMessage}</div>;

    return (
      <div>
        {/* Your JSX here */}
      </div>
    );
  }
}
```

### Props Interface — Always Be Explicit

```typescript
// IMyFeatureProps.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyFeatureProps {
  context: WebPartContext;      // Strongly type context — NOT 'any'
  listName: string;
  showExportButton: boolean;
  newFormUrl: string;
  myhttpclient: any;            // HttpClient — 'any' is acceptable here
}
```

---

## 11. Building Forms — State, Validation & Handlers

### Form State Pattern

For class components, define a state interface and use `setState`. For functional components, use the `handleInputChange` utility from `FormHandlers.ts`.

#### Class Component Form State

```typescript
interface IFormState {
  title: string;
  status: string;
  amount: number;
  selectedDate: Date;
  assignedTo: any[];   // People Picker value
  attachments: File[];
}

interface IValidationErrors {
  validationErrorTitle?: string;
  validationErrorStatus?: string;
  validationErrorAmount?: string;
  validationErrorDate?: string;
  validationErrorAssignedTo?: string;
}
```

#### Functional Component — Using FormHandlers

```typescript
import { handleInputChange, IFormState, IValidationState }
  from '../../../framework/Services/ListOperations/FormHandlers';

const [formState, setFormState] = React.useState<IFormState>({
  title: '', status: 'Draft', amount: 0
});
const [validationState, setValidationState] = React.useState<IValidationState>({});

// handleInputChange:
// 1. Updates the named field in formState
// 2. Clears the validation error for that field automatically
const onFieldChange = (fieldName: string, value: any) => {
  handleInputChange(fieldName, value, setFormState, setValidationState);
};

// Error key pattern: "validationError" + CapitalizedFieldName
// Field 'title' → key 'validationErrorTitle'
// Field 'amount' → key 'validationErrorAmount'
```

### Complete Validation Reference

```typescript
const _validator = new ValidationService();

// Text / single-line field
const [err, msg] = _validator.isTextFieldEmpty(value, "");

// Rich text (strips HTML tags before checking)
const [err, msg] = _validator.isRichTextFieldEmpty(htmlValue, "");

// Date picker (null = empty)
const [err, msg] = _validator.isDatePickerEmpty(dateValue, "");

// Number field (0 or null = invalid)
const [err, msg] = _validator.isNumberFieldEmpty(numValue, "");

// People Picker (must have at least 1 person)
const [err, msg] = _validator.isPeoplePickerEmpty(peopleArray);

// Single dropdown/choice
const [err, msg] = _validator.isDropdownEmpty(value, "");

// Multi-select dropdown
const [err, msg] = _validator.isMultiSelectDropdownEmpty(valueArray, "");

// Checkbox (must be checked = true)
const [err, msg] = _validator.isCheckboxUnchecked(boolValue, "");

// File attachments
const [err, msg] = _validator.isAttachmentEmpty(fileArray);

// File picker control
const [err, msg] = _validator.isFilePickerEmpty(value, "");
```

### Displaying Validation Errors (Fluent UI)

```tsx
import { TextField, Dropdown, DatePicker } from '@fluentui/react';

<TextField
  label="Title"
  required
  value={this.state.formState.title}
  onChange={(_, val) => this.setState(prev => ({
    formState: { ...prev.formState, title: val || '' },
    validationErrors: { ...prev.validationErrors, validationErrorTitle: '' }
  }))}
  errorMessage={this.state.validationErrors.validationErrorTitle}
/>
```

### People Picker — Critical: Save as User ID

```typescript
// Step 1: User selects a person (PeoplePicker control gives you email/loginName)
const onPersonSelected = async (selectedPeople: any[]) => {
  // Step 2: Resolve to SharePoint User ID
  const email = selectedPeople[0]?.secondaryText; // email from control
  const userId = await this._listService.getUserId(email);

  // Step 3: Save to list using the ID field (field name + "Id" suffix)
  await this._listService.addItemsToList(Constant.listName, {
    Title: "My Item",
    AssignedToId: userId        // NOT AssignedTo — must be the "Id" field
  });
};
```

---

## 12. Building Dashboards — Tables, Filters & Exports

### The Dashboard Pattern

The Common Platform uses `material-react-table` for dashboard views. The recommended pattern separates column configuration from data fetching:

```
RootComponent (Dashboard.tsx)
    ├── Initializes services (componentDidMount)
    ├── Defines column configuration (array of column objects)
    ├── Defines which SP columns to fetch (listColumns array)
    └── Passes all of this to → MainDashboard component
```

### Column Configuration

```typescript
const columns: any[] = [
  // Simple text column
  {
    accessorKey: 'Title',
    header: 'Title',
    size: 200,
    filterFn: (row, columnId, filterValue) => {
      const cellValue = row.getValue(columnId);
      if (cellValue == null) return false;
      return cellValue.toString().toLowerCase().includes(filterValue.toLowerCase());
    }
  },

  // Number column with formatting
  {
    accessorKey: 'Amount',
    header: 'Amount',
    size: 150,
    // Custom filter that matches the formatted value (with commas)
    filterFn: (row, columnId, filterValue) =>
      this._utilities.numberColumnFilterForDashboard(
        row, columnId, filterValue, this._utilities.formatNumberWithComma
      ),
    Cell: (cell: any) =>
      cell.row.original.Amount == null
        ? ""
        : this._utilities.formatNumberWithComma(cell.row.original.Amount)
  },

  // Date column — correctly handles timezone offset
  {
    accessorKey: 'StartDate',
    header: 'Start Date',
    size: 150,
    sortingFn: 'datetime',
    accessorFn: (row: any) => row?.StartDate ?? '',
    filterFn: (term, columnName, searchText) =>
      this._utilities.handleDateFilterForDashboard(term, columnName, searchText, "StartDate"),
    Cell: ({ cell }: { cell: any }) =>
      this._utilities.formatDateWithOffset(cell.getValue(), this.offset)
  },

  // Status column with color coding
  {
    accessorKey: 'Status',
    header: 'Status',
    Cell: (cell: any) => (
      <div style={{
        color: cell.renderedCellValue === 'Approved' ? 'green'
             : cell.renderedCellValue === 'Pending'  ? '#f59e0b'
             : cell.renderedCellValue === 'Rejected' ? 'red'
             : 'black'
      }}>
        {cell.renderedCellValue}
      </div>
    )
  },

  // Hyperlink column (for Edit links)
  {
    accessorKey: 'EditLink',
    header: 'Edit',
    enableColumnFilter: false,
    enableSorting: false,
    enableGrouping: false,
    size: 80,
    Cell: ({ cell }: any) => {
      const link = cell.getValue();
      return link?.Url
        ? <a href={link.Url} target="_blank" data-interception="off" rel="noopener noreferrer">
            {link.Description || "Edit"}
          </a>
        : null;
    }
  }
];
```

### List Columns to Fetch

```typescript
// These are the SharePoint internal column names to fetch from the list
// Always define this separately — it controls what data gets loaded
const listColumns = [
  "Id", "Title", "Status", "Amount", "StartDate", "EditLink",
  "Author/Title"  // Lookup — needs to also be in expand
];
```

### Export Configuration

```typescript
// Separate column config for export — may include/exclude different columns
const exportColumns = [
  { accessorKey: 'Id', header: 'ID' },
  { accessorKey: 'Title', header: 'Title' },
  { accessorKey: 'Status', header: 'Status' },
  { accessorKey: 'Amount', header: 'Amount' },
  { accessorKey: 'StartDate', header: 'Start Date' }
];
```

### CSV Export

```typescript
import { CSVExportService } from '../../../framework/Utilities/CSVExportService';

private exportToCSV(): void {
  const csvService = new CSVExportService(Constant.exportFileName);
  // Transform data if needed before export
  const exportData = this.state.items.map(item => ({
    ID: item.Id,
    Title: item.Title,
    Status: item.Status,
    Amount: this._utilities.formatNumberCustom(item.Amount, 2)
  }));
  csvService.export(exportData);
}
```

### Excel File Reading (from SharePoint Library)

```typescript
import { ExcelUtilities } from '../../../framework/Utilities/ExcelUtilities';

private async readExcelFromSharePoint(): Promise<void> {
  const _excel = new ExcelUtilities();
  const serverRelativePath = "/sites/MySite/Shared Documents/data.xlsx";
  const data = await _excel.getExcelFileDataFromPath(serverRelativePath, this._listService);
  // data is an array of row objects keyed by header names
  this.setState({ importedData: data });
}
```

---

## 13. Error Handling & Logging — Never Skip This

### Why Logging Matters in SPFx

In production SharePoint environments, developers rarely have direct access to logs or debug output. Without structured error logging to a SharePoint list, bugs are impossible to diagnose.

### The Logging List

Create a `LogHistory` list in SharePoint with these columns:

| Column Name | Internal Name | Type |
|---|---|---|
| Title | `Title` | Single line of text |
| Page Name | `gen_PageName` | Single line of text |
| Message | `gen_Message` | Single line of text |
| Exception Message | `gen_ExceptionMessage` | Multiple lines of text |
| Exception Date | `gen_ExceptionMessageDate` | Date and Time |

### The Log Method (from PnpListOperationService)

```typescript
public async addItemsToLogList(
    listName: string,        // Always: Constant.LogHistoryList
    functionName: string,    // The function where the error occurred
    pageName: string,        // The component/page (be specific!)
    message: string,         // Human-readable description
    exceptionmsg: string,    // The actual error
    errorDate: Date          // Timestamp
): Promise<any> {
    // Implementation writes to the LogHistory SP list
}
```

### Standard Error Handling Pattern

```typescript
// ✅ Correct pattern — async/await with structured logging
public async myOperation(): Promise<IMyListItem[]> {
    try {
        const items = await this._sp.web.lists
            .getByTitle(Constant.listName)
            .items
            .select("Id", "Title")
            .filter("Status eq 'Active'")();
        return items;
    } catch (error) {
        console.error("MyComponent.myOperation: Failed", error);

        // Log to SharePoint — always do this in service layer
        await this.addItemsToLogList(
            Constant.LogHistoryList,
            "myOperation",              // Function name — exact
            "MyFeatureMain",            // Component name — be specific, not "PageName"
            "Error fetching active items from list",
            String(error),
            new Date()
        );

        // Decide: rethrow or return fallback?
        // For data fetches: return fallback to not break the UI
        return [];
        // For critical mutations (create/update/delete): rethrow so UI can handle
        // throw error;
    }
}
```

### When to Rethrow vs Return Fallback

| Operation | Action | Why |
|---|---|---|
| Fetching list items | Return `[]` | Dashboard stays visible, user sees empty state |
| Fetching a single item | Return `null` | Let UI show "not found" gracefully |
| Adding/Updating/Deleting | `throw error` | Caller (component) must handle and show error toast |
| File upload | `throw error` | User must know the upload failed |
| Non-critical utility | Return `''` or `0` | Default value is acceptable |

### Show Errors to Users (Toast Notifications)

```typescript
import { toast } from 'react-toastify';

private async onSave(): Promise<void> {
    try {
        await this._listService.addItemsToList(Constant.listName, item);
        toast.success("Item saved successfully!");
    } catch (error) {
        toast.error("Failed to save. Please try again.");
        console.error(error);
    }
}
```

---

## 14. Performance Optimization Rules

### Rule 1: Always Use `select()` — Never Fetch All Columns

Every list item has 40+ system columns. Fetching all of them wastes bandwidth and slows the webpart.

```typescript
// Wrong — fetches all 40+ columns
const items = await this._sp.web.lists.getByTitle("MyList").items();

// Correct — only fetch what the UI needs
const items = await this._sp.web.lists.getByTitle("MyList").items
    .select("Id", "Title", "Status", "Amount", "Author/Title")
    .expand("Author")();
```

### Rule 2: Use Batch for Multiple Writes

```typescript
// Wrong — 100 separate HTTP requests
for (const item of items) {
    await this._listService.addItemsToList(Constant.listName, item); // N requests
}

// Correct — 1 batched HTTP request
await this._listService.addMultipleItemsToList(Constant.listName, items);
```

### Rule 3: Use Async Iterator for Lists That May Exceed 5000 Items

```typescript
// Wrong — fails silently if list has > 5000 items (returns only first 5000)
const items = await this._sp.web.lists.getByTitle("MyList").items.top(5000)();

// Correct — iterates through all pages automatically
for await (const page of this._sp.web.lists.getByTitle("MyList").items.top(5000)) {
    allItems.push(...page);
}
```

### Rule 4: Initialize Services Once in `componentDidMount`

```typescript
// Wrong — creates new service instance on every render (called many times)
public render(): React.ReactElement {
    const service = new IPnpListOperationService(); // BAD — new instance per render!
    // ...
}

// Correct — created once, reused for the component lifetime
private _listService: IPnpListOperationService = null;

async componentDidMount(): Promise<void> {
    this._listService = new IPnpListOperationService();
    await this._listService.Init(this.props.context);
}
```

### Rule 5: Guard Against Null Services in `render()`

On the very first render, `componentDidMount` has not completed yet. Services are `null`. Any call to a service inside `render()` will crash.

```typescript
// Wrong — crashes on first render
Cell: (cell: any) => this._utilities.formatNumberWithComma(cell.value)

// Correct — null guard
Cell: (cell: any) =>
    this._utilities
        ? this._utilities.formatNumberWithComma(cell.value)
        : cell.value
```

### Rule 6: Use OData Filter Server-Side, Not Client-Side

```typescript
// Wrong — fetches ALL items, then filters in JavaScript
const allItems = await _listService.GetAllItemsFromList("MyList", "", [], []);
const activeItems = allItems.filter(i => i.Status === 'Active'); // Client-side filter

// Correct — filter is applied on the server, less data transferred
const activeItems = await _listService.GetAllItemsFromList(
    "MyList",
    "Status eq 'Active'",  // OData filter — runs on server
    ["Id", "Title"],
    []
);
```

### Rule 7: Lazy Load Heavy Libraries

For large dependencies like `jsPDF`, `xlsx`, or `html2canvas`, consider lazy loading:

```typescript
// For PDF — only import when user clicks "Print"
private async onPrint(): Promise<void> {
    // html2canvas and jsPDF are heavy — only load when needed
    await this._utilities.printAsPdf();
}
```

### Rule 8: Prefer `async/await` Consistently

Mixed `async/await` + `.then().catch()` chains create maintenance issues. Pick one style per project and stick to it.

```typescript
// Consistent async/await (recommended)
try {
    const result = await someOperation();
} catch (error) {
    // handle
}

// Avoid mixing styles in the same codebase
someOperation()
    .then(result => { /* ... */ })
    .catch(error => { /* ... */ }); // Only if async/await isn't possible
```

---

## 15. Common SPFx Pitfalls to Avoid

### Pitfall 1: Using `context` in the Constructor

```typescript
// WRONG — context is not available in constructor
constructor(props: IMyProps) {
    super(props);
    this._sp = getSP(this.props.context); // context is null here!
}

// CORRECT — use componentDidMount
async componentDidMount(): Promise<void> {
    this._listService = new IPnpListOperationService();
    await this._listService.Init(this.props.context); // context is available here
}
```

### Pitfall 2: People Picker Saves Email Instead of User ID

```typescript
// WRONG — throws "Invalid data" error from SharePoint
await _listService.addItemsToList("MyList", {
    AssignedTo: "john@company.com" // ❌ SP expects a user ID integer
});

// CORRECT — resolve ID first
const userId = await _listService.getUserId("john@company.com");
await _listService.addItemsToList("MyList", {
    AssignedToId: userId // ✅ Integer User ID
});
```

### Pitfall 3: Graph Without Declared Permissions

```typescript
// This will fail at runtime with 401 Unauthorized if you haven't declared:
// { "resource": "Microsoft Graph", "scope": "User.Read.All" }
// in package-solution.json AND approved it in SP Admin Center
const users = await graphService.getAllUsers(this.props.context);
```

**Fix:** Always add permissions to `package-solution.json` → deploy → approve in Admin Center before writing Graph code.

### Pitfall 4: Not Handling Graph Pagination

```typescript
// WRONG — only returns first 100 users; misses the rest
const response = await client.api("/users").get();
const users = response.value;

// CORRECT — handles all pages via @odata.nextLink
let users = [];
let nextLink = "/users?$top=999";
while (nextLink) {
    const response = await client.api(nextLink).get();
    users = [...users, ...response.value];
    nextLink = response["@odata.nextLink"] || null;
}
```

### Pitfall 5: Hardcoding List Names / Library Names

```typescript
// WRONG — if the list is renamed, you have to find and change every occurrence
const items = await _listService.GetAllItemsFromList("CapexCommonTest", ...);

// CORRECT — change in one place, all references update
// In Constant.ts: public static listName: string = "CapexCommonTest";
const items = await _listService.GetAllItemsFromList(Constant.listName, ...);
```

### Pitfall 6: Missing `onDispose()` — Memory Leaks

```typescript
// WRONG — React component stays mounted after webpart is removed
// Missing onDispose()

// CORRECT — always unmount
protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
}
```

### Pitfall 7: Folder Path vs Library Name Confusion

Different PnP methods use different path formats:

```typescript
// Methods that use LIBRARY TITLE (display name)
this._sp.web.lists.getByTitle("Shared Documents").rootFolder.folders()

// Methods that use SERVER-RELATIVE PATH (full URL path)
this._sp.web.getFolderByServerRelativePath("/sites/MySite/Shared Documents/MyFolder").files()
```

**Rule of thumb:** Methods starting with `getByTitle(libraryName)` use the display name. Methods with `getFolderByServerRelativePath` or `getFileByServerRelativePath` need the full server-relative path.

### Pitfall 8: Not Updating `dataVersion` After Changing WebPart Properties

If you add or remove properties from `IMyWebPartProps`, increment the `dataVersion`:

```typescript
// After adding new properties
protected get dataVersion(): Version {
    return Version.parse('1.1'); // Increment from 1.0
}
```

Failing to do this can cause property pane migration issues for existing webpart placements.

### Pitfall 9: Using Wrong Graph Client Version

```typescript
// WRONG — deprecated, may cause auth failures
const client = await context.msGraphClientFactory.getClient("1");

// CORRECT — always use version "3"
const client = await context.msGraphClientFactory.getClient("3");
```

### Pitfall 10: Treating SP Date Strings as Local Dates

SharePoint stores dates in UTC. If you display them without a timezone offset correction, dates may appear one day off.

```typescript
// WRONG — ignores timezone, may show wrong date
const displayDate = new Date(item.StartDate).toLocaleDateString();

// CORRECT — apply timezone offset from getSPCurrentTimeZone()
const offset = await this._utilities.getSPCurrentTimeZone();
const displayDate = this._utilities.formatDateWithOffset(item.StartDate, offset);
```

## 16. Build, Bundle & Deploy Workflow

Understanding how to test and compile your SPFx solution is critical. Below are the steps for local development testing and production packaging under both toolchains.

### Local Development & Testing

Before deploying, run a local server instance to test changes with hot reloading.

#### 1. Start the Local Server
Depending on your project's version, execute the appropriate command in your terminal:

*   **Gulp Toolchain (SPFx < v1.22):**
    ```bash
    gulp serve
    ```
*   **Heft Toolchain (SPFx v1.22+):**
    ```bash
    # Starts the local development task (runs compilation, watches files, starts local web server)
    heft start
    ```

#### 2. Previewing Your WebPart
*   **SharePoint Hosted Workbench (Recommended):**  
    Modern SPFx development requires testing directly inside a real SharePoint site. Navigate to the hosted workbench:
    ```
    https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/15/workbench.aspx
    ```
*   **Testing on Modern Pages:**  
    You can also test the local webpart on a live modern page by adding the query string parameter `?debugManifestsFile=https://localhost:4321/temp/manifests.js` to the page URL and selecting **Load debug scripts** when prompted.

---

### Production Build Sequence

To compile, optimize, and bundle your webpart assets for production deployment, run the compile and package commands sequentially:

#### Using Gulp Toolchain (Legacy / v1.20.0):
```bash
# 1. Clean previous build files and folder structures
gulp clean

# 2. Compile, minimize resources, and build production manifests
gulp bundle --ship

# 3. Create the deployment .sppkg package file
gulp package-solution --ship
```

#### Using Heft Toolchain (Modern / v1.22.0+):
```bash
# 1. Clean previous build output
npx heft clean

# 2. Build and bundle production-optimized assets
npx heft build --production

# 3. Package the solution into a deployable *.sppkg archive
npx heft package-solution --production
```

Output: The packaged file is created at `sharepoint/solution/<project-name>.sppkg`.

---

### Deployment Checklist

**Step 1 — App Catalog Deployment**
- [ ] Upload `.sppkg` to **Tenant App Catalog** (or site collection app catalog)
- [ ] Click **Deploy** — choose "Make this solution available to all sites" for tenant-wide
- [ ] Confirm the solution version is updated (matches `package-solution.json`)

**Step 2 — Graph API Permissions**
- [ ] Go to **SharePoint Admin Center → Advanced → API access**
- [ ] Review all pending permission requests from your solution
- [ ] Click **Approve** for each required permission
- [ ] Verify no permissions are stuck in "Pending"

**Step 3 — Site Deployment**
- [ ] Go to target site → **Site contents → Add an app**
- [ ] Find your solution and add it to the site
- [ ] Edit the page, add the webpart
- [ ] Configure the Property Pane settings (list names, URLs, toggles)
- [ ] Publish the page

**Step 4 — Post-Deployment Testing**
- [ ] Test form submission (add item to list)
- [ ] Test form validation (required fields)
- [ ] Test data loading in dashboard
- [ ] Test export (CSV/Excel)
- [ ] Test filters and sorting in dashboard
- [ ] Check browser console — zero errors should appear
- [ ] Test in different browsers (Edge, Chrome)

### Versioning Strategy

Every time you deploy an update:
1. Increment `version` in `package-solution.json` (e.g., `1.0.0.0` → `1.0.0.1` for patch, `1.1.0.0` for feature)
2. If properties schema changed, increment `dataVersion` in the WebPart file
3. Document what changed in a `CHANGELOG.md` or commit message

---

## 17. Code Quality Checklist

Use this checklist before every pull request or deployment:

### Architecture
- [ ] No business logic in `WebPart.ts` — only React element creation
- [ ] All list/library names referenced via `Constant.*` — no hardcoded strings
- [ ] All new reusable functions added to `framework/` — not duplicated in each webpart
- [ ] Services initialized in `componentDidMount` — not in constructor or `render()`

### TypeScript Quality
- [ ] State interfaces defined — no `this.state = {} as any`
- [ ] Props interfaces defined — avoid `any` for typed props like `context`
- [ ] Data model interfaces defined for list items — avoid `any[]` for item arrays
- [ ] No `debugger` statements in code
- [ ] No commented-out blocks of dead code left in production files

### Data Operations
- [ ] All list reads use `select()` — no fetching of all columns
- [ ] Lists that could exceed 5000 items use `GetAllItemsFromList` (async iterator)
- [ ] Multiple writes use batching methods
- [ ] People Picker saves use resolved User IDs — not emails
- [ ] Deletes use `recycle()` — not hard delete
- [ ] OData filters applied server-side — not JavaScript array filters post-fetch

### Error Handling
- [ ] Every `try/catch` block includes a `console.error` call
- [ ] Every caught error in service layer logs to `addItemsToLogList`
- [ ] The `pageName` parameter in `addItemsToLogList` is the actual component name
- [ ] User-facing errors shown via toast notification — not just console logs
- [ ] Loading and error states defined in component state

### Graph API
- [ ] Required permissions declared in `package-solution.json`
- [ ] Graph client always uses version `"3"`
- [ ] Pagination (`@odata.nextLink`) handled in all Graph calls that return lists
- [ ] Principle of least privilege — only necessary scopes requested

### Memory & Performance
- [ ] `onDispose()` calls `ReactDom.unmountComponentAtNode`
- [ ] Null guards in `render()` for services initialized in `componentDidMount`
- [ ] No service instantiation inside `render()`
- [ ] Heavy libraries (jsPDF, xlsx) only used when actually triggered by user action

---

## 18. Quick Reference

### Service Initialization Decision Table

| Service | Needs `Init(context)`? | Context Passed |
|---|---|---|
| `IPnpListOperationService` | YES | Once in `Init()` — stored as `_sp` |
| `ICommonUtilities` | YES | Once in `Init()` — stored as `_sp` |
| `IGraphAPIOperationService` | NO | Per method call |
| `ValidationService` | NO | Not needed |
| `CSVExportService` | NO | Filename in constructor |
| `ExcelUtilities` | NO | `_listService` passed per method |

### Import Paths (from `components/` folder inside a webpart)

```typescript
import { IPnpListOperationService }
  from '../../../framework/Services/ListOperations/PnpListOperationService';

import { IGraphAPIOperationService }
  from '../../../framework/Services/ListOperations/GraphAPIOperationService';

import { ValidationService }
  from '../../../framework/Services/ListOperations/ValidationService';

import { handleInputChange, IFormState, IValidationState }
  from '../../../framework/Services/ListOperations/FormHandlers';

import { ICommonUtilities }
  from '../../../framework/Utilities/CommonUtilities';

import { ExcelUtilities }
  from '../../../framework/Utilities/ExcelUtilities';

import { CSVExportService }
  from '../../../framework/Utilities/CSVExportService';

import { Constant }
  from '../../../framework/Constants/Constant';
```

### CRUD Quick Patterns

```typescript
// READ (all with filter)
const items = await _listService.GetAllItemsFromList(
  Constant.listName,
  "Status eq 'Active'",              // OData filter
  ["Id", "Title", "Status"],         // Select
  []                                 // Expand
);

// READ (by ID)
const item = await _listService.getItemById(
  Constant.listName, itemId, ["Id", "Title"], []
);

// CREATE (single)
await _listService.addItemsToList(Constant.listName, { Title: "New", Status: "Draft" });

// CREATE (batch)
await _listService.addMultipleItemsToList(Constant.listName, arrayOfItems);

// UPDATE (single)
await _listService.updateItemInList(Constant.listName, itemId, { Status: "Approved" });

// UPDATE (batch)
await _listService.updateListItemsInBatch(Constant.listName, [
  { id: 1, fields: { Status: "Approved" } },
  { id: 2, fields: { Status: "Rejected" } }
]);

// DELETE (recycle bin — recoverable)
await _listService.deleteItemFromList(Constant.listName, itemId);
```

### OData Filter Quick Reference

```
Equals:           "Title eq 'Value'"
Not equals:       "Title ne 'Value'"
Contains:         "substringof('Value', Title)"
Starts with:      "startswith(Title, 'Value')"
Greater than:     "Amount gt 1000"
Less than:        "Amount lt 1000"
AND:              "Status eq 'Active' and Amount gt 100"
OR:               "Status eq 'Active' or Status eq 'Pending'"
Null check:       "Title ne null"
Boolean:          "IsApproved eq 1"
Date comparison:  "StartDate gt '2026-01-01T00:00:00Z'"
Lookup field:     "Category/Title eq 'Finance'"
```

### Common Validation Patterns

```typescript
const v = new ValidationService();

// Returns [isError: boolean, message: string]
v.isTextFieldEmpty(value, "")
v.isRichTextFieldEmpty(htmlValue, "")
v.isDatePickerEmpty(dateValue, "")
v.isNumberFieldEmpty(numValue, "")      // 0 or null = error
v.isPeoplePickerEmpty(peopleArray)      // empty array = error
v.isDropdownEmpty(value, "")
v.isMultiSelectDropdownEmpty(arr, "")
v.isCheckboxUnchecked(boolValue, "")
v.isAttachmentEmpty(fileArray)
v.isFilePickerEmpty(value, "")
```

### Build Commands

```bash
gulp serve                   # Local dev with hot reload
heft start                   # Alternative local dev (heft)
gulp clean                   # Clean build output
gulp bundle --ship           # Production build (minified)
gulp package-solution --ship # Create .sppkg package
```

### Utility Method Reference

```typescript
// Date formatting
utils.formatDateTime(date, "DD/MM/YYYY")           // moment.js format
utils.formatDateWithOffset(dateStr, offset, 'en-GB') // With TZ offset
utils.isDateValid(value)                             // boolean
await utils.getSPCurrentTimeZone()                   // Returns offset number

// Number formatting
utils.formatNumberCustom(value, 2)                   // "1,234,567.89"
utils.formatNumberWithComma(value, 0)                // "1,234,568"
utils.isNumberNullOrInvalid(value)                   // true if 0/null/NaN

// String utilities
utils.IsStrNullOrEmpty(str)                          // true if null/whitespace
utils.extractTextFromHTML("<p>Hello</p>")            // "Hello"

// Array utilities
utils.sortByKey(array, 'fieldName', 'asc'|'desc')   // Returns sorted copy
utils.groupBy(array, 'propertyName')                 // Returns grouped object

// URL
utils.getParameterByName("paramName", url)           // Query string value

// Email (needs Init)
await utils.sendCustomEmail(to[], cc[], subject, htmlBody)

// External API
await utils.callCustomApi(httpClient, SPHttpClient, data, apiUrl)

// Power Automate
await _listService.callPowerAutomateFlow(flowUrl, payload)

// PDF (looks for id="formToPrint" in DOM)
await utils.printAsPdf()
```

---

*Guide Version: 2.0 | Last Updated: August 2026*  
*Reference Implementation: Common Platform (SPFx 1.20.0)*  
*Enterprise Development Standards & Architecture*
