export type ComplianceForType = 'Business' | 'Employee' | 'Project' | 'Subcontractor' | 'Worker' | '';

export interface IComplianceItem {
  Id: number;
  ComplianceFor: ComplianceForType;
  ComplianceTypeId: number | null;
  ComplianceTypeName: string;
  BusinessId: number | null;
  BusinessName: string;
  EmployeeId: number | null;
  EmployeeName: string;
  ProjectId: number | null;
  ProjectTitle: string;
  SubcontractorId: number | null;
  SubcontractorName: string;
  WorkerId: number | null;
  WorkerName: string;
  IsBooking: boolean;
  BookingDate: string | null;
  BookedWith: string;
  DocumentNumber: string;
  IssuingAuthority: string;
  IssueDate: string | null;
  RenewalNotRequired: boolean;
  ExpiryDate: string | null;
  Notes: string;
  HasAttachments: boolean;
  MainBusinessProfileId?: number | null;

  // Visual/Calculated properties for UBuild Grid
  EntityActive?: string;
  Status?: string;
  Expired?: string;
  DaysRemaining?: string;

  // Audit / extra fields for export
  CreatedBy?: string;
  DateEntered?: string;
  Position?: string;       // Employee or Worker position from concatenated string
  BusinessProfile?: string; // Resolved Business Profile name
  TypeDescription?: string; // Compliance Type description if available
}

export interface ILookupOption {
  key: number;
  text: string;
}

export interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface ILookupSets {
  complianceTypes: ILookupOption[];
  businesses: ILookupOption[];
  employees: ILookupOption[];
  projects: ILookupOption[];
  subcontractors: ILookupOption[];
  workers: ILookupOption[];
}
