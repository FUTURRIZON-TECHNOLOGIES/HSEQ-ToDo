export const WORKSITE_HAC_SWMS_MODULE_ID = 'WorksiteHacSwms';
export const WORKSITE_HAC_SWMS_MODULE_LABEL = 'Worksite HAC & SWMS';
export const WORKSITE_HAC_SWMS_MODULE_ICON = 'ClipboardList';
export const WORKSITE_HAC_SWMS_MODULE_GROUP = 'HSEQ' as const;

export interface IWorksiteHacSwmsListConfig {
    hacListTitle: string;
    swmsListTitle: string;
    projectsListTitle: string;
    hazardChecklistListTitle: string;
    preWorkChecklistListTitle: string;
    emergencyResponseListTitle: string;
    contactsListTitle: string;
    visitorsListTitle: string;
    workPartyMembersListTitle: string;
}

export const defaultWorksiteHacSwmsListConfig: IWorksiteHacSwmsListConfig = {
    hacListTitle: '',
    swmsListTitle: '',
    projectsListTitle: '',
    hazardChecklistListTitle: '',
    preWorkChecklistListTitle: '',
    emergencyResponseListTitle: '',
    contactsListTitle: '',
    visitorsListTitle: '',
    workPartyMembersListTitle: ''
};
