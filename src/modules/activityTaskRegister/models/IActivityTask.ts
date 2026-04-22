export interface ISPUser {
    Id: number;
    Title: string;
    Email?: string;
}

export interface ISPLookup {
    Id: number;
    Title: string;
}

export interface IActivityTaskItem {
    Id: number;
    Title?: string; // Often Task or similar is used as Title
    Activity?: ISPLookup;
    ActivityId?: number;
    WorkZone?: ISPLookup;
    WorkZoneId?: number;
    Task?: string;
    ResponsiblePersons?: ISPUser;
    ResponsiblePersonsId?: number;
    Consequence?: string;
    Likelihood?: string;
    RiskRanking?: string;
    RevisedConsequence?: string;
    RevisedLikelihood?: string;
    RevisedRanking?: string;
    HighRiskWork?: boolean;
    Active?: boolean;
    
    // Virtual fields for grid display
    ActivityValue?: string;
    WorkZoneValue?: string;
    ResponsiblePersonsValue?: string;
    
    Created?: string;
    Modified?: string;
    Author?: ISPUser;
    Editor?: ISPUser;
}

export interface IActivityTaskHazard {
    Id: number;
    Title: string; // Hazard description
    ActivityTaskRegisterId: number;
    HazardId?: number; // Lookup to Hazard Type if applicable
    Hazard?: ISPLookup;
}
