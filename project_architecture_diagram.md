# SPFx Employee Leave Management Architecture & Workflow

This diagram illustrates how data flows through the application, starting from the user interface down to the SharePoint list, utilizing the newly established architectural layers.

## 1. High-Level Architecture & Data Flow

```mermaid
graph TD
    %% Define styles for different layers
    classDef ui fill:#e1f0f2,stroke:#03787c,stroke-width:2px,color:#000
    classDef feature fill:#fff4ce,stroke:#8a6d00,stroke-width:2px,color:#000
    classDef framework fill:#dff6dd,stroke:#107c10,stroke-width:2px,color:#000
    classDef sp fill:#0078d4,stroke:#005a9e,stroke-width:2px,color:#fff

    %% SharePoint Environment
    subgraph SharePoint ["SharePoint Online Environment"]
        SPList["📋 Employee Leave Requests (List)"]:::sp
        LogList["📜 LogHistory (List)"]:::sp
    end

    %% WebPart Entry
    WebPart["🧩 CrudOperationWebPart\n(Entry Point)"]:::ui

    %% React UI Layer (Feature)
    subgraph ReactUI ["React UI Layer (components/)"]
        CrudOperation["🖥️ CrudOperation\n(State Controller)"]:::ui
        LeaveRequestList["📊 LeaveRequestList\n(Table & Filters)"]:::ui
        LeaveRequestForm["📝 LeaveRequestForm\n(Inputs & Validation)"]:::ui
        PeoplePicker["👤 PeoplePickerField"]:::ui
    end

    %% Feature Services Layer
    subgraph FeatureService ["Feature Services (services/)"]
        LeaveReqService["⚙️ LeaveRequestService\n(Business Logic)"]:::feature
    end

    %% Shared Framework Layer
    subgraph SharedFramework ["Shared Framework (framework/)"]
        Constants["📌 Constant.ts"]:::framework
        PnPConfig["🔌 pnpjsConfig.ts\n(SPFI Singleton)"]:::framework
        Utilities["🧰 Utilities.ts"]:::framework
        Logger["🛡️ LoggingService.ts"]:::framework
    end

    %% Wiring it all together
    WebPart -->|Initializes & Passes Context| CrudOperation
    
    CrudOperation -->|Passes Items| LeaveRequestList
    CrudOperation -->|Handles Submit| LeaveRequestForm
    LeaveRequestForm -->|Selects User| PeoplePicker
    
    CrudOperation -->|Calls CRUD Methods| LeaveReqService
    
    LeaveReqService -->|Reads List Names| Constants
    LeaveReqService -->|Gets Instance| PnPConfig
    LeaveReqService -->|Logs Errors| Logger
    
    LeaveRequestForm -.->|Validates Dates| Utilities
    LeaveRequestList -.->|Formats Dates| Utilities
    
    PnPConfig ==>|REST API| SPList
    Logger -.->|Writes Errors| LogList
```

## 2. User Action Workflow (Example: Creating a Leave Request)

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant UI as LeaveRequestForm (React)
    participant Ctrl as CrudOperation (React)
    participant Svc as LeaveRequestService (Service)
    participant FW as pnpjsConfig (Framework)
    participant SP as SharePoint List

    User->>UI: Fills out Leave details & clicks "Submit"
    UI->>UI: Validates mandatory fields & date ranges
    
    alt Validation Fails
        UI-->>User: Shows inline error messages
    else Validation Succeeds
        UI->>Ctrl: handleSubmit(inputData)
    end
    
    Ctrl->>Svc: service.create(inputData)
    
    Svc->>FW: getSP(context)
    FW-->>Svc: Returns PnP SPFI Singleton
    
    Svc->>SP: REST POST /items (Status: 'Pending')
    
    alt API Call Succeeds
        SP-->>Svc: Success Response (201 Created)
        Svc-->>Ctrl: Resolves Promise
        Ctrl-->>User: Closes Panel & Shows "Success" MessageBar
        Ctrl->>Svc: loadItems() (Refreshes Data)
    else API Call Fails
        SP-->>Svc: Error Response (500/400)
        Svc->>LoggingService: logError(context, error)
        LoggingService->>SharePoint LogHistory: Writes error to log
        Svc-->>Ctrl: Throws Error
        Ctrl-->>User: Shows "Error" MessageBar
    end
```
