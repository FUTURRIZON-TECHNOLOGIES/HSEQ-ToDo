/**
 * Mapping configuration for dynamic fields based on 'Regarding' selection.
 * Key: The display value in the 'Regarding' choice dropdown.
 * Value: The SharePoint internal name of the corresponding text column.
 */
export const REGARDING_DYNAMIC_FIELDS: { [key: string]: string } = {
    "Audit & Inspection": "AuditInspection",
    "Clients": "Clients",
    "Compliance Register": "ComplianceRegister",
    "Employee": "Employee",
    "Incident": "Incident",
    "Leads": "Leads",
    "Meetings": "Meetings",
    "Project": "Project",
    "Proposal": "Proposal",
    "Subcontractor": "Subcontractor",
    "Subcontractor Employee": "SubcontractorEmployee",
    "Submission": "Submission",
    "Training & Induction": "TrainingInduction",
    "Vehicle & Plant": "VehiclePlant"
};

/**
 * Returns a human-readable label for the dynamic field based on its internal name.
 */
export const getRegardingFieldLabel = (internalName: string): string => {
    switch (internalName) {
        case "AuditInspection":        return "Audit & Inspection";
        case "Clients":                return "Clients";
        case "ComplianceRegister":     return "Compliance Register";
        case "Employee":               return "Employee";
        case "Incident":               return "Incident";
        case "Leads":                  return "Leads";
        case "Meetings":               return "Meetings";
        case "Project":                return "Project";
        case "Proposal":               return "Proposal";
        case "Subcontractor":          return "Subcontractor";
        case "SubcontractorEmployee":  return "Subcontractor Employee";
        case "Submission":             return "Submission";
        case "TrainingInduction":      return "Training & Induction";
        case "VehiclePlant":           return "Vehicle & Plant";
        default:                       return "Details";
    }
};
