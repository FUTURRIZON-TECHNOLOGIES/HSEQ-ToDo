export interface IToDoItem {
    Id?: number;
    Title: string; // Subject
    Description?: string;
    Comment?: string;

    // Lookups (Objects for display)
    Status?: { Id: number; Title: string; Name?: string };
    Category?: { Id: number; Title: string; Name?: string };
    Classification?: { Id: number; Title: string; Name?: string };
    Priority?: { Id: number; Title: string; Name?: string };
    TrainingInduction?: { Id: number; Title: string };

    // People (Objects for display)
    TaskOwner?: { Id: number; Title: string; EMail: string };
    AssigneeInternal?: { Id: number; Title: string; EMail: string };
    AssigneeExternal?: { Id: number; Title: string; EMail: string };
    CreatedByUser?: { Id: number; Title: string; EMail: string };
    UpdatedByUser?: { Id: number; Title: string; EMail: string };
    PersonName?: { Id: number; Title: string; EMail: string };

    // Standard Fields
    Regarding?: string; // Choice
    DueDate?: string;
    StartDate?: string;
    EstimationDuration?: string;
    ActualDuration?: string;
    CompletionDate?: string;
    CompletedPercent?: number; // Completed %
    EmailNotifications?: boolean;
    Resolution?: string;

    // Metadata
    Created?: string;
    Author?: { Title: string };
    Modified?: string;
    Editor?: { Title: string };
    Attachments?: boolean;
    AttachmentFiles?: IAttachment[];

    // Field IDs for Saving
    StatusId?: number;
    CategoryId?: number;
    ClassificationId?: number;
    PriorityId?: number;
    TaskOwnerId?: number;
    AssigneeInternalId?: number;
    AssigneeExternalId?: number;
    PersonNameId?: number;
    TrainingInductionId?: number;

    // Dynamic 'Regarding' Fields
    AuditInspection?: string;
    Clients?: string;
    ComplianceRegister?: string;
    Employee?: string;
    Incident?: string;
    Leads?: string;
    Meetings?: string;
    Project?: string;
    Proposal?: string;
    Subcontractor?: string;
    SubcontractorEmployee?: string;
    Submission?: string;
    VehiclePlant?: string;
}

export interface IAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

export interface ILookupOption {
    Id: number;
    Title: string;
}
