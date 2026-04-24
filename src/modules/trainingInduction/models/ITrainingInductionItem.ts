export interface ILookupOption {
    Id: number;
    Title: string;
    Name?: string;
}

export interface IPerson {
    Id: number;
    Title: string;
    EMail: string;
}

export interface IAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

export interface ITrainingInductionItem {
    Id?: number;
    Title?: string; // Type (Choice/Title)
    Type?: string;
    TrainingFor?: string;
    TrainingType?: string;
    Participant?: ILookupOption; // Lookup(Contacts -> Employee Name)
    ScheduledDate?: string;
    Status?: string; // (Scheduled, In Progress, Completed)

    // Outcome Section
    InvitationSection?: string; // Multiple line text
    InductionLink?: {
        Description: string;
        Url: string;
    };
    ParticipantsStatus?: string; // (Active, Inactive)
    InvitationStatus?: string;
    SendInvitation?: boolean;

    // Internal Section
    BusinessProfile?: ILookupOption;
    Manager?: ILookupOption;
    Supervisors?: ILookupOption;
    Coordinator?: ILookupOption;

    // UI Helpers
    Company?: ILookupOption;
    Project?: ILookupOption;
    CompletionDate?: string;
    Overdue?: string;
    InductionFor?: string;

    // Metadata
    Author?: IPerson;
    Editor?: IPerson;
    Created?: string;
    Modified?: string;
    AttachmentFiles?: IAttachment[];
}
