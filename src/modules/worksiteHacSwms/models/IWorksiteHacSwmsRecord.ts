export interface ISwmsReference {
    id: string;
    title: string;
}

export interface IHacContact {
    role: string;
    name: string;
    number: string;
    company: string;
}

export interface IHacHazardRow {
    hazard: string;
    rating: string;
    controlMethods: string;
    residualRating: string;
}

export interface IHacTimelineEntry {
    id: string;
    avatarText: string;
    name: string;
    dateText: string;
    description: string;
    imageUrl?: string;
}

export interface IWorksiteHacSwmsRecord {
    id: number;
    number: string;
    date: string;
    project: string;
    scopeOfWorks: string;
    workAddresses: string;
    weatherCondition: string;
    status: string;
    submittedBy: string;
    supervisor: string;
    geoLocation: string;
    emergencyMusterLocation: string;
    firstAidKitLocation: string;
    nearestMedicalCentre: string;
    nearestHospital: string;
    swmsUsed: ISwmsReference[];
    contacts: IHacContact[];
    hazards: IHacHazardRow[];
    timeline: IHacTimelineEntry[];
}
