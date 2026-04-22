import { IWorksiteHacSwmsRecord } from '../models/IWorksiteHacSwmsRecord';

export const worksiteHacSwmsRecords: IWorksiteHacSwmsRecord[] = [
    {
        id: 11712,
        number: '11712',
        date: '2026-03-25',
        project: '32281 - Thornton - Training and Assessment',
        scopeOfWorks: 'Level 2 Overhead and Underground',
        workAddresses: 'Lot 8C/13 Hartley Drive, Thornton NSW 2322',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Adam Razbusek',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [{ id: '015', title: 'IAC WHS Training Facility Safety' }],
        contacts: [
            { role: 'Supervisor/Coordinator', name: 'Adam Razbusek', number: '0433 722 974', company: 'ASP Assist Group' },
            { role: 'Nominated First Aider', name: '', number: '', company: '' },
            { role: 'Emergency Contact', name: '', number: '', company: '' },
            { role: 'Client Contact Present', name: '', number: '', company: '' }
        ],
        hazards: [
            {
                hazard: 'Emergency Situations',
                rating: 'Extreme',
                controlMethods: 'Follow emergency procedures, keep exits clear, move calmly to the assembly point, and maintain required PPE.',
                residualRating: 'Low'
            },
            {
                hazard: 'Entering and Moving Around Training Facilities',
                rating: 'High',
                controlMethods: 'Keep walkways clear, follow trainer instructions, and always walk while moving through the worksite.',
                residualRating: 'Low'
            },
            {
                hazard: 'Lifting and Moving Training Equipment',
                rating: 'High',
                controlMethods: 'Use suitable lifting equipment, avoid lifting over people, and request help for awkward loads.',
                residualRating: 'Low'
            }
        ],
        timeline: [
            { id: '1', avatarText: 'AA', name: 'Auditor ASP Assist Group', dateText: 'April 07, 05:33 PM', description: 'Daily HAC Detail Updated' },
            { id: '2', avatarText: 'AR', name: 'Adam Razbusek', dateText: 'March 25, 07:04 AM', description: 'New Worksite HAC or Pre-Start Created' }
        ]
    },
    {
        id: 11689,
        number: '11689',
        date: '2025-08-12',
        project: '32280 - Blacktown - Training and Assessment',
        scopeOfWorks: 'UEECD0007 UEECD0019 UEECD0051',
        workAddresses: 'Blacktown Training Facility',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Training Coordinator',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [{ id: '018', title: 'Electrical Training Controls' }],
        contacts: [],
        hazards: [],
        timeline: []
    },
    {
        id: 11638,
        number: '11638',
        date: '2024-06-12',
        project: '32285 - MRTU - Training and Assessment',
        scopeOfWorks: 'Training',
        workAddresses: 'MRTU Training Site',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Training Coordinator',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [],
        contacts: [],
        hazards: [],
        timeline: []
    },
    {
        id: 11659,
        number: '11659',
        date: '2024-11-20',
        project: '13000 - BT - General',
        scopeOfWorks: 'Testing',
        workAddresses: 'BT General Site',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Worksite Coordinator',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [],
        contacts: [],
        hazards: [],
        timeline: []
    },
    {
        id: 11698,
        number: '11698',
        date: '2025-09-23',
        project: '32280 - Blacktown - Training and Assessment',
        scopeOfWorks: 'Example of SWMS used for the refresher course',
        workAddresses: 'Blacktown Training Facility',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Training Coordinator',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [{ id: '022', title: 'Refresher Course Safety' }],
        contacts: [],
        hazards: [],
        timeline: []
    },
    {
        id: 11714,
        number: '11714',
        date: '2026-04-08',
        project: '32284 - Port Macquarie - Training and Assessment',
        scopeOfWorks: 'testing',
        workAddresses: 'Port Macquarie Training Site',
        weatherCondition: '',
        status: 'In Progress',
        submittedBy: '',
        supervisor: 'Worksite Coordinator',
        geoLocation: '(No Location)',
        emergencyMusterLocation: '',
        firstAidKitLocation: '',
        nearestMedicalCentre: '',
        nearestHospital: '',
        swmsUsed: [],
        contacts: [],
        hazards: [],
        timeline: []
    }
];

export const projectOptions: string[] = Array.from(new Set(worksiteHacSwmsRecords.map(item => item.project)));
export const statusOptions: string[] = ['Draft', 'In Progress', 'Completed', 'Closed'];
