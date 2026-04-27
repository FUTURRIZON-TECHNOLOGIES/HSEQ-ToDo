import * as React from 'react';
import {
    TextField,
    IDropdownOption,
    ComboBox,
    DatePicker,
    DefaultButton,
    PrimaryButton,
    Icon,
    MessageBar,
    MessageBarType,
    Pivot,
    PivotItem,
    Modal,
    IconButton,
    IContextualMenuProps
} from '@fluentui/react';
import { ExportService } from '../../webparts/swf/services/ExportService';
import { ITrainingInductionItem } from './models/ITrainingInductionItem';
import { TrainingInductionService } from './services/TrainingInductionService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import styles from './TrainingInduction.module.scss';
import ToDoModule from '../todo/ToDoModule';

export interface ITrainingInductionFormProps {
    item: ITrainingInductionItem | null;
    spService: TrainingInductionService;
    context: WebPartContext;
    onClose: () => void;
    onSave: (payload: any, mode: 'stay' | 'close' | 'new') => Promise<void>;
    preLoadedLookups?: { participants: IDropdownOption[], businessProfiles: IDropdownOption[], employees: IDropdownOption[], trainingTypes: IDropdownOption[] };
    defaultTab?: 'detail' | 'attachments' | 'actions';
}

export interface IDocument {
    Id: number;
    Title: string;
    DocumentType: { Id: number, Name: string };
    FileName: string;
    Description: string;
    Created: string;
    Author: { Title: string };
    ServerRelativeUrl: string;
}

interface IUploadItem {
    file: File;
    title: string;
    docType: string;
    description: string;
}

const ToDoModuleAny = ToDoModule as React.FC<any>;

const DOCUMENT_LIBRARY_NAME = "Training  Inductions Documents";

const FILE_ICON_MAP: Record<string, string> = {
    pdf: 'PDF', doc: 'WordDocument', docx: 'WordDocument',
    xls: 'ExcelDocument', xlsx: 'ExcelDocument',
    ppt: 'PowerPointDocument', pptx: 'PowerPointDocument',
    jpg: 'Photo', jpeg: 'Photo', png: 'Photo', gif: 'Photo', bmp: 'Photo', webp: 'Photo',
    zip: 'ZipFolder', rar: 'ZipFolder', '7z': 'ZipFolder',
    txt: 'TextDocument', csv: 'TextDocument', msg: 'Mail'
};

const getFileIcon = (fileName: string): string => {
    const ext = (fileName || '').split('.').pop()?.toLowerCase() || '';
    return FILE_ICON_MAP[ext] || 'Document';
};

const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const TYPE_OPTIONS: IDropdownOption[] = [
    "Supply Workforce - WHS Management Plan v1",
    "Supply Workforce - New Subcontractor Induction Questionnaire (v3)",
    "Supply Workforce - New Subcontractor Induction Questionnaire (v2)",
    "Supply Workforce - Environmental Management Plan v1",
    "Supply Workforce - Quality Management Plan v1",
    "Angle Grinder Safety", "Asbestos Awareness", "Ausgrid OHSP - Safety Observer training", "Behavioural Safety",
    "Bloodborne Pathogens – Managing the Risk", "Building and Office Evacuation", "Bullying in the Workplace",
    "Burns", "Chain of Responsibility", "Chemical Safety", "Competent observer essential energy", "Conflict of Interest",
    "Contractor Induction", "Correct Mask application and use", "CPR - Cardiopulmonary Resuscitation",
    "Drone Excluded category safety rules", "Drugs and Alcohol at Work", "Environmental Awareness",
    "FOR-2112-Intellectual Property and Confidentiality Agreement", "New Employee Induction Questionnaire (Version 1)",
    "Probity Plan - Sydney Metro", "Risk Register Training for Senior Managers", "Safe Manual Handling",
    "SMR Induction", "The Safe Use of Ladders", "Ultegra Company Induction", "Work Related Stress", "Working from Home Fundamentals"
].map(opt => ({ key: opt, text: opt }));

const STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'Scheduled', text: 'Scheduled' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Complete', text: 'Complete' }
];

const PARTICIPANT_STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'Active', text: 'Active' },
    { key: 'Inactive', text: 'Inactive' }
];

const TrainingInductionForm: React.FC<ITrainingInductionFormProps> = (props) => {
    const [formData, setFormData] = React.useState<Partial<ITrainingInductionItem>>(props.item || {
        Status: 'Scheduled',
        Title: 'Supply Workforce - WHS Management Plan v1'
    });
    const [saving, setSaving] = React.useState(false);
    const [error, setError] = React.useState<string | null>(null);
    const [expandedSections, setExpandedSections] = React.useState<Record<string, boolean>>({
        general: true, outcome: true, internal: true, timeline: true
    });
    const [activeTab, setActiveTab] = React.useState<string>(props.defaultTab || 'detail');

    React.useEffect(() => {
        setActiveTab(props.defaultTab || 'detail');
    }, [props.item?.Id, props.defaultTab]);

    React.useEffect(() => {
        if (activeTab === 'attachments') {
            fetchLibraryDocs();
        }
    }, [props.item?.Id, activeTab]);

    React.useEffect(() => {
        if (props.item) {
            setFormData(props.item);
        }
    }, [props.item]);
    const [lookupOptions, setLookupOptions] = React.useState<{
        participants: IDropdownOption[],
        businessProfiles: IDropdownOption[],
        employees: IDropdownOption[],
        trainingTypes: IDropdownOption[],
        documentTypes: IDropdownOption[]
    }>({
        participants: props.preLoadedLookups?.participants || [],
        businessProfiles: props.preLoadedLookups?.businessProfiles || [],
        employees: props.preLoadedLookups?.employees || [],
        trainingTypes: props.preLoadedLookups?.trainingTypes || [],
        documentTypes: []
    });

    const [libraryDocs, setLibraryDocs] = React.useState<IDocument[]>([]);
    const [isUploadModalOpen, setIsUploadModalOpen] = React.useState(false);
    const [uploadQueue, setUploadQueue] = React.useState<IUploadItem[]>([]);
    const [selectedDocIds, setSelectedDocIds] = React.useState<number[]>([]);
    const [columnFilters, setColumnFilters] = React.useState<Record<string, string>>({});
    const [isDragging, setIsDragging] = React.useState(false);
    const selectAllRef = React.useRef<HTMLInputElement>(null);

    const filteredDocs = React.useMemo(() => {
        return libraryDocs.filter(doc => {
            const lc = (s: string) => s.toLowerCase();
            return (
                (!columnFilters.title || lc(doc.Title || '').includes(lc(columnFilters.title))) &&
                (!columnFilters.docType || lc(doc.DocumentType?.Name || '').includes(lc(columnFilters.docType))) &&
                (!columnFilters.fileName || lc(doc.FileName || '').includes(lc(columnFilters.fileName))) &&
                (!columnFilters.description || lc(doc.Description || '').includes(lc(columnFilters.description))) &&
                (!columnFilters.uploadedBy || lc(doc.Author?.Title || '').includes(lc(columnFilters.uploadedBy)))
            );
        });
    }, [libraryDocs, columnFilters]);

    const isAllSelected = filteredDocs.length > 0 && filteredDocs.every(d => selectedDocIds.includes(d.Id));
    const isSomeSelected = !isAllSelected && filteredDocs.some(d => selectedDocIds.includes(d.Id));

    React.useEffect(() => {
        if (selectAllRef.current) selectAllRef.current.indeterminate = isSomeSelected;
    }, [isSomeSelected]);

    React.useEffect(() => { setSelectedDocIds([]); }, [libraryDocs]);

    const handleSelectAll = (): void => {
        setSelectedDocIds(isAllSelected ? [] : filteredDocs.map(d => d.Id));
    };

    const handleSelectOne = (id: number): void => {
        setSelectedDocIds(prev => prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]);
    };

    React.useEffect(() => {
        if (!props.preLoadedLookups) {
            const load = async () => {
                try {
                    const [pts, bps, emps, ttypes, dtypes] = await Promise.all([
                        props.spService.getLookupOptions('Contacts', 'Employee Name'),
                        props.spService.getLookupOptions('BusinessProfiles', 'BusinessProfile'),
                        props.spService.getLookupOptions('Employees', 'Employee Name'),
                        props.spService.getLookupOptions('TrainingTypes', 'Title'),
                        props.spService.getLookupOptions('Document Types', 'Name')
                    ]);
                    setLookupOptions({
                        participants: pts.map(p => ({ key: p.Id, text: p.Title })),
                        businessProfiles: bps.map(b => ({ key: b.Id, text: b.Title })),
                        employees: emps.map(e => ({ key: e.Id, text: e.Title })),
                        trainingTypes: ttypes.map(t => ({ key: t.Title, text: t.Title })),
                        documentTypes: dtypes.map(d => ({ key: d.Id, text: d.Title }))
                    });
                } catch (e) { console.error('Lookup load failed', e); }
            };
            load();
        } else {
            const pre = props.preLoadedLookups!;
            setLookupOptions(prev => ({
                participants: pre.participants,
                businessProfiles: pre.businessProfiles,
                employees: pre.employees,
                trainingTypes: pre.trainingTypes,
                documentTypes: prev.documentTypes
            }));
            props.spService.getLookupOptions('Document Types', 'Name')
                .then(dtypes => setLookupOptions(prev => ({ ...prev, documentTypes: dtypes.map(d => ({ key: d.Id, text: d.Title })) })))
                .catch(e => console.error('Document types load failed', e));
        }
    }, [props.preLoadedLookups]);

    const [isEditModalOpen, setIsEditModalOpen] = React.useState(false);
    const [editingDoc, setEditingDoc] = React.useState<IDocument | null>(null);
    const actionTrainingInductionValue = React.useMemo(() => {
        if (!props.item?.Id) return "";
        const typeValue =
            typeof formData.Type === 'object'
                ? ((formData.Type as any)?.Title || (formData.Type as any)?.Name || "")
                : (formData.Type || "");

        return `${props.item.Id} - ${String(typeValue || "").trim()}`;
    }, [props.item?.Id, formData.Type]);

    const fetchLibraryDocs = async (): Promise<void> => {
        if (props.item?.Id) {
            console.log(`[InductionForm] !!! FETCH INITIATED !!! Target ID: ${props.item.Id}, Library: ${DOCUMENT_LIBRARY_NAME}`);
            try {
                const docs = await props.spService.getLibraryDocuments(DOCUMENT_LIBRARY_NAME, props.item.Id);
                console.log(`[InductionForm] !!! FETCH SUCCESS !!! Found ${docs.length} documents for ID ${props.item.Id}`);
                if (docs.length === 0) {
                    console.warn(`[InductionForm] WARNING: No documents found for ID ${props.item.Id}. Ensure 'RecordID' column in library has THIS value.`);
                } else {
                    console.table(docs.map(d => ({ ID: d.Id, Title: d.Title, FileName: d.FileName, Type: d.DocumentType?.Name })));
                }
                setLibraryDocs(docs);
            } catch (err) {
                console.error("[InductionForm] Critical error during fetchLibraryDocs:", err);
            }
        } else {
            console.warn("[InductionForm] Cannot fetch docs: props.item.Id is missing.");
        }
    };

    const handleFieldChange = (field: keyof ITrainingInductionItem, value: any): void => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    const handleSave = async (mode: 'stay' | 'close' | 'new'): Promise<void> => {
        setSaving(true);
        setError(null);
        try {
            const payload: any = {
                Title: formData.Title,
                Type: formData.Type,
                TrainingFor: formData.TrainingFor,
                Status: formData.Status,
                ScheduledDate: formData.ScheduledDate,
                InvitationStatus: formData.InvitationStatus,
                ParticipantsStatus: formData.ParticipantsStatus,
                SendInvitation: formData.SendInvitation
            };
            if (formData.Participant?.Id) {
                payload.ParticipantId = formData.Participant.Id;
                payload.ParticipantsId = formData.Participant.Id;
            }
            if (formData.BusinessProfile?.Id) {
                payload.BusinessProfileId = formData.BusinessProfile.Id;
                payload.Business_x0020_ProfileId = formData.BusinessProfile.Id;
            }
            if (formData.Manager?.Id) payload.ManagerId = formData.Manager.Id;
            if (formData.Supervisors?.Id) payload.SupervisorsId = formData.Supervisors.Id;
            if (formData.Coordinator?.Id) payload.CoordinatorId = formData.Coordinator.Id;

            await props.onSave(payload, mode);
        } catch (e) {
            setError(e.message || 'Error occurred while saving.');
        } finally {
            setSaving(false);
        }
    };

    const handleDownloadCertificate = (): void => {
        const id = props.item?.Id || 'N/A';
        const name = (formData.Participant?.Title || 'Participant Name').trim();

        // Prioritize Type field for the title as requested
        const typeValue = typeof formData.Type === 'string' ? formData.Type : '';
        const trainingTitle = (typeValue || formData.Title || `Training #${id}`).trim();

        const rawDate = formData.CompletionDate || formData.ScheduledDate;
        const dateStr = rawDate
            ? new Date(rawDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })
            : 'Date';

        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const certBgUrl: string = require('./certificateBackground.jpg');

        const img = new Image();
        img.onload = (): void => {
            const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
            const pageWidth = doc.internal.pageSize.getWidth();   // 841.89 pt
            const pageHeight = doc.internal.pageSize.getHeight(); // 595.28 pt
            const cx = pageWidth / 2;

            // Colors
            const GOLD: [number, number, number] = [198, 156, 28];
            const NAVY: [number, number, number] = [34, 76, 115];
            const DARK_TEXT: [number, number, number] = [33, 37, 41];

            // 1. Background Image (Borders/Graphics only)
            doc.addImage(img, 'JPEG', 0, 0, pageWidth, pageHeight);

            // 2. Header: "CERTIFICATE"
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(54);
            doc.setTextColor(GOLD[0], GOLD[1], GOLD[2]);
            doc.text('CERTIFICATE', cx, 140, { align: 'center' });

            // 3. Sub-Header: "of Completion"
            doc.setFontSize(22);
            doc.setTextColor(NAVY[0], NAVY[1], NAVY[2]);
            doc.text('of Completion', cx, 175, { align: 'center' });

            // 4. Awarded To Line
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(18);
            doc.setTextColor(DARK_TEXT[0], DARK_TEXT[1], DARK_TEXT[2]);
            doc.text('This Certificate is awarded to:', cx, 245, { align: 'center' });

            // 5. Name (Dynamic)
            let nameSize = 58;
            if (name.length > 20) nameSize = 44;
            if (name.length > 30) nameSize = 34;

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(nameSize);
            doc.setTextColor(NAVY[0], NAVY[1], NAVY[2]);
            doc.text(name, cx, 315, { align: 'center' });

            // 6. Training Description
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(16);
            doc.setTextColor(DARK_TEXT[0], DARK_TEXT[1], DARK_TEXT[2]);
            doc.text("for successfully completing the training of", cx, 395, { align: 'center' });

            // 7. Training Title (Dynamic, Bold)
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(20);
            doc.setTextColor(0, 0, 0);
            const titleLines = doc.splitTextToSize(trainingTitle, 620);
            doc.text(titleLines, cx, 425, { align: 'center' });

            const titleHeight = (titleLines.length) * 24;
            const dateY = 425 + titleHeight;

            // 8. Date (Dynamic)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(16);
            doc.setTextColor(DARK_TEXT[0], DARK_TEXT[1], DARK_TEXT[2]);
            doc.text(`on ${dateStr}`, cx, dateY, { align: 'center' });

            // 9. Certificate ID (Subtle)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8);
            doc.setTextColor(150, 150, 150);
            doc.text(`Certificate No: ${id}`, pageWidth - 40, pageHeight - 25, { align: 'right' });

            doc.save(`Certificate_${id}.pdf`);
        };
        img.src = certBgUrl;
    };

    const handleShowResult = (): void => {
        console.log("[InductionForm] Generating Result PDF...");
        try {
            const doc = new jsPDF({
                orientation: 'p',
                unit: 'pt',
                format: 'a4'
            });

            const id = props.item?.Id || 'New';

            const name = formData.Participant?.Title || 'Participant Name';
            const company = formData.Company?.Title || formData.BusinessProfile?.Title || '';
            const type = formData.Title || (typeof formData.Type === 'string' ? formData.Type : '') || 'Induction Name';
            const rawDate = formData.CompletionDate || formData.ScheduledDate;
            const dateStr = rawDate ? new Date(rawDate).toLocaleDateString('en-GB', {
                weekday: 'long',
                day: 'numeric',
                month: 'long',
                year: 'numeric'
            }) : '';

            // --- PAGE 1 ---
            
            // Logo Placeholder
            doc.setFillColor(20, 50, 90);
            doc.rect(50, 40, 80, 45, 'F');
            doc.setTextColor(255, 255, 255);
            doc.setFontSize(10);
            doc.setFont('helvetica', 'bold');
            doc.text('SUPPLY', 60, 58);
            doc.text('WORKFORCE', 54, 72);
            
            doc.setTextColor(38, 71, 114);
            doc.setFontSize(26);
            doc.setFont('helvetica', 'bold');
            doc.text(`Training/Induction # ${id}`, 545, 65, { align: 'right' });

            // Detail Header
            const startY = 110;
            doc.setFillColor(38, 71, 114);
            doc.rect(50, startY, 495, 22, 'F');
            doc.setTextColor(255, 255, 255);
            doc.setFontSize(11);
            doc.text('TRAINING/INDUCTION DETAIL', 55, startY + 15);

            // Detail Table
            autoTable(doc, {
                startY: startY + 22,
                theme: 'grid',
                headStyles: { fillColor: [245, 245, 245], textColor: [0,0,0], fontStyle: 'bold', lineWidth: 0.5 },
                bodyStyles: { textColor: [0, 0, 0], lineWidth: 0.5, fontSize: 10 },
                margin: { left: 50, right: 50 },
                columnStyles: {
                    0: { cellWidth: 160, fillColor: [250, 250, 250], fontStyle: 'bold' },
                    1: { cellWidth: 'auto' }
                },
                body: [
                    ['Record Number', id.toString()],
                    ['Participant', name],
                    ['Company', company],
                    ['Type', type],
                    ['Training/Induction For', formData.TrainingFor || 'General'],
                    ['Project (optional)', formData.Project?.Title || ''],
                    ['Completion Date', dateStr],
                    ['Status', formData.Status || 'Complete'],
                    ['Participant\'s Signature', { content: '', styles: { minCellHeight: 40 } }]
                ]
            });

            // Mock Signature logic (Robust check)
            const lat = (doc as any).lastAutoTable;
            const finalY = lat ? lat.finalY : startY + 250;
            
            doc.setDrawColor(0,0,0);
            doc.setLineWidth(1);
            doc.moveTo(320, finalY - 15);
            (doc as any).bezierCurveTo(340, finalY - 35, 360, finalY + 5, 380, finalY - 20);
            doc.stroke();

            // Result Title
            doc.setFillColor(38, 71, 114);
            doc.rect(50, finalY + 15, 495, 22, 'F');
            doc.setTextColor(255, 255, 255);
            doc.text('TRAINING/INDUCTION RESULT', 55, finalY + 30);

            // Result Sections
            autoTable(doc, {
                startY: finalY + 37,
                theme: 'grid',
                headStyles: { fillColor: [240, 240, 240], textColor: [0,0,0], fontStyle: 'bold' },
                bodyStyles: { fontSize: 9 },
                margin: { left: 50, right: 50 },
                columnStyles: { 0: { cellWidth: 340 }, 1: { cellWidth: 'auto' } },
                head: [['Presentation', '']],
                body: [['I have READ AND UNDERSTOOD the Supply Workforce New Subcontractor Induction Presentation.', 'Yes']]
            });

            const nextY = (doc as any).lastAutoTable?.finalY || finalY + 100;
            autoTable(doc, {
                startY: nextY + 5,
                theme: 'grid',
                headStyles: { fillColor: [240, 240, 240], textColor: [0,0,0], fontStyle: 'bold' },
                bodyStyles: { fontSize: 9 },
                margin: { left: 50, right: 50 },
                columnStyles: { 0: { cellWidth: 340 }, 1: { cellWidth: 'auto' } },
                head: [['Questions', '']],
                body: [
                    ['Select four (4) of Supply Workforce Core Values:', 'Honesty, Teamwork, Safety First, Respect'],
                    ['What must be undertaken prior to starting any work?', 'Risk assessment undertaken'],
                    ['Who is responsible for completing a risk assessment?', 'All members of the work party'],
                    ['What are three (3) risk analysis tools we can use to determine if a task is safe to perform?', 'Risk matrix, Hierarchy of controls, SWMS'],
                    ['What is the correct order to effectively perform a risk assessment?', 'Assess -> Identify -> Control -> Monitor'],
                    ['Identify five (5) typical hazards to be controlled in metering works:', 'Animals, Live electricity, Asbestos, Customers, Radiation'],
                    ['What is the purpose of a Safe Work Method Statement (SWMS)?', 'To identify high risk tasks and associated minimum controls...'],
                    ['Select six (6) attributes that promote good customer relations:', 'Respect, Cleanliness, Greeting, Notification, Groomed, Identification'],
                    ['What do we do if we identify an unidentified risk?', 'Stop work, inform team, assess controls, update risk assessment'],
                    ['Who is responsible for your safety at work? (Select 3)', 'All members of the work party, Supervisors, Me'],
                    ['How do I report a safety concern? (Select 2)', 'Report to my supervisor'],
                    ['You will be subjected to random drug and alcohol tests.', 'True'],
                    ['When must pre-start checks be done on vehicles?', 'Daily before commencing work'],
                    ['What clothing is required near electricity network?', 'Arc rated clothing, Neck to wrists to ankles']
                ]
            });

            const addFooter = (currDoc: jsPDF, page: number, total: number) => {
                currDoc.setTextColor(150, 150, 150);
                currDoc.setFontSize(8);
                currDoc.text(`Training & Induction printout. Printed on: ${new Date().toLocaleString()}. Documents may not be current.`, 50, 815);
                currDoc.setFillColor(20, 50, 80);
                currDoc.rect(530, 805, 15, 15, 'F');
                currDoc.setTextColor(255, 255, 255);
                currDoc.text(`${page}/${total}`, 537.5, 815.5, { align: 'center' });
            };
            addFooter(doc, 1, 2);

            // PAGE 2
            doc.addPage();
            addFooter(doc, 2, 2);
            doc.setTextColor(38, 71, 114);
            doc.setFontSize(26);
            doc.text(`Training/Induction # ${id}`, 545, 65, { align: 'right' });

            const page2StartY = 100;
            autoTable(doc, {
                startY: page2StartY,
                theme: 'grid',
                headStyles: { fillColor: [240, 240, 240], textColor: [0,0,0], fontStyle: 'bold' },
                bodyStyles: { fontSize: 9 },
                margin: { left: 50, right: 50 },
                columnStyles: { 0: { cellWidth: 340 }, 1: { cellWidth: 'auto' } },
                head: [['Questions (Continued)', '']],
                body: [
                    ['How do I obtain PPE or safety equipment?', 'Request from supervisor'],
                    ['What action if a customer is aggressive/abusive?', 'Sympathise and refer to Supply Workforce supervisor'],
                    ['What if you encounter a dog requiring access?', 'Do not enter unless safe and dog is secured by owner'],
                    ['What if you find friable asbestos swarf?', 'Apply PPE as per SWMS and clean board prior to works'],
                    ['Should employees be aware of social media activity?', 'True'],
                    ['Briefed on behavioural expectations?', 'Yes'],
                    ['Other type of business Bluecurrent has?', 'Gas'],
                    ['Types of operations in Wellington:', 'Distribution network provider'],
                    ['Role of AEMO?', 'To manage Australia\'s electricity and gas systems'],
                    ['Doorstep Protocol - advice to customer:', 'Technician should give no advice on configuration/tariff'],
                    ['If an incident occurs, notify who?', 'Emergency services, SWF Supervisor, Bluecurrent, Safework']
                ]
            });

            const plansY = (doc as any).lastAutoTable?.finalY || page2StartY + 300;
            autoTable(doc, {
                startY: plansY + 10,
                theme: 'grid',
                margin: { left: 50, right: 50 },
                bodyStyles: { fontSize: 9 },
                columnStyles: { 0: { cellWidth: 340, fontStyle: 'bold' }, 1: { cellWidth: 'auto' } },
                body: [
                    ['Supply Workforce - Quality Management Plan v1', 'Yes'],
                    ['Supply Workforce - WHS Management Plan v1', 'Yes'],
                    ['Supply Workforce - Environmental Management Plan v1', 'Yes']
                ]
            });

            doc.save(`InductionResult_${id}.pdf`);
        } catch (error) {
            console.error("PDF Generation failed:", error);
            alert("Error generating PDF: " + (error.message || "Unknown error"));
        }
    };

    const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>): void => {
        if (!e.target.files) return;
        const files = Array.from(e.target.files);
        const newItems = files.map(f => ({
            file: f,
            title: f.name.split('.').slice(0, -1).join('.'),
            docType: "",
            description: ""
        }));
        setUploadQueue([...uploadQueue, ...newItems]);
    };

    const saveUploadedDocs = async (): Promise<void> => {
        if (!props.item?.Id) return;
        setSaving(true);
        try {
            for (const item of uploadQueue) {
                await props.spService.uploadLibraryDocument(DOCUMENT_LIBRARY_NAME, props.item.Id, item.file);
            }
            alert("Documents uploaded successfully.");
            setUploadQueue([]);
            setIsUploadModalOpen(false);
            fetchLibraryDocs();
        } catch (e) {
            console.error("[InductionForm] Document upload failed:", e);
            setError("Failed to upload documents: " + (e.message || "Please check if the 'RecordID' column exists in the library."));
        } finally {
            setSaving(false);
        }
    };

    const handleUpdateMetadata = async (): Promise<void> => {
        if (!editingDoc) return;
        setSaving(true);
        try {
            await props.spService.updateLibraryDocumentMetadata(DOCUMENT_LIBRARY_NAME, editingDoc.Id, {
                Title: editingDoc.Title,
                DocumentType: editingDoc.DocumentType.Id,
                Description: editingDoc.Description
            });
            setIsEditModalOpen(false);
            await fetchLibraryDocs();
        } catch (e: any) {
            setError("Failed to update document metadata: " + (e?.message || String(e)));
        } finally {
            setSaving(false);
        }
    };

    const handleExportLibraryDocs = (type: 'excel' | 'csv'): void => {
        if (libraryDocs.length === 0) return;
        const data = libraryDocs.map((doc, idx) => ({
            '#': idx + 1,
            'Title': doc.Title,
            'Document Type': doc.DocumentType.Name,
            'File Name': doc.FileName,
            'Description': doc.Description,
            'Date Uploaded': new Date(doc.Created).toLocaleString(),
            'Uploaded By': doc.Author.Title
        }));
        const fileName = `Attachments_${props.item?.Id || 'Record'}_${new Date().getTime()}`;
        if (type === 'excel') {
            ExportService.exportToExcel(data as any, fileName);
        } else {
            ExportService.exportToCSV(data as any, fileName);
        }
    };

    const exportMenuProps: IContextualMenuProps = {
        items: [
            {
                key: 'excel',
                text: 'Export to Excel',
                iconProps: { iconName: 'ExcelDocument' },
                onClick: () => handleExportLibraryDocs('excel')
            },
            {
                key: 'csv',
                text: 'Export to CSV',
                iconProps: { iconName: 'TextDocument' },
                onClick: () => handleExportLibraryDocs('csv')
            }
        ],
    };

    const deleteSelectedDocs = async (): Promise<void> => {
        if (selectedDocIds.length === 0) return;
        if (!confirm(`Are you sure you want to delete ${selectedDocIds.length} document(s)?`)) return;

        setSaving(true);
        try {
            for (const id of selectedDocIds) {
                await props.spService.deleteLibraryDocument(DOCUMENT_LIBRARY_NAME, id);
            }
            alert("Documents deleted successfully.");
            setSelectedDocIds([]);
            fetchLibraryDocs();
        } catch (e) {
            setError("Failed to delete documents.");
        } finally {
            setSaving(false);
        }
    };









    const renderTimeline = (): JSX.Element => {
        const item = props.item;
        if (!item) return <p className={styles.timelineEmpty}>Timeline will populate after saving.</p>;

        return (
            <div className={styles.timeline}>
                {item.Modified && (
                    <div className={styles.timelineEvent}>
                        <div className={styles.timelineDot} />
                        <div className={styles.timelineContent}>
                            <strong>{item.Editor?.Title || 'User'}</strong>
                            <span className={styles.timelineDate}>
                                {new Date(item.Modified).toLocaleString()}
                            </span>
                            <span className={styles.timelineDesc}>Induction Record Updated</span>
                        </div>
                    </div>
                )}
                {item.Created && (
                    <div className={styles.timelineEvent}>
                        <div className={`${styles.timelineDot} ${styles.timelineDotLast}`} />
                        <div className={styles.timelineContent}>
                            <strong>{item.Author?.Title || 'User'}</strong>
                            <span className={styles.timelineDate}>
                                {new Date(item.Created).toLocaleString()}
                            </span>
                            <span className={styles.timelineDesc}>New Induction Record Created</span>
                        </div>
                    </div>
                )}
            </div>
        );
    };

    const toggleSection = (section: string) => {
        setExpandedSections(prev => ({ ...prev, [section]: !prev[section] }));
    };

    const Section: React.FC<{ title: string; id: string; children: React.ReactNode }> = ({ title, id, children }) => {
        const isExpanded = expandedSections[id];
        return (
            <div className={styles.section} style={{ marginBottom: '15px' }}>
                <div
                    className={styles.sectionHeader}
                    onClick={() => toggleSection(id)}
                    style={{ cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                >
                    <Icon
                        iconName="ChevronRight"
                        className={styles.chevron}
                        style={{
                            transform: isExpanded ? 'rotate(90deg)' : 'rotate(0deg)',
                            transition: 'transform 0.3s ease'
                        }}
                    />
                    <span className={styles.sectionTitle}>{title}</span>
                </div>
                {isExpanded && <div className={styles.sectionBody}>{children}</div>}
            </div>
        );
    };

    return (
        <div className={styles.todoForm}>
            <div className={styles.toolbar}>
                <div className={styles.formTitle}>
                    <Icon iconName="Education" className={styles.formTitleIcon} />
                    <span>{props.item ? `Training/Induction: ${props.item.Id}` : 'New Training/Induction'}</span>
                </div>
                <div className={styles.toolbarActions}>
                    <PrimaryButton
                        className={styles.btnSave}
                        iconProps={{ iconName: 'Save' }}
                        text={saving ? 'Saving…' : 'Save'}
                        disabled={saving}
                        onClick={() => handleSave('stay')}
                    />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'SaveAs' }}
                        text="Save & New"
                        disabled={saving}
                        onClick={() => handleSave('new')}
                    />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'SaveAndClose' }}
                        text="Save & Close"
                        disabled={saving}
                        onClick={() => handleSave('close')}
                    />
                    <div className={styles.toolbarDivider} />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'Refresh' }}
                        text="Refresh"
                        disabled={saving}
                        onClick={() => {
                            // Re-load initial data if needed or just trigger a state refresh
                            window.location.reload();
                        }}
                    />
                    <DefaultButton
                        className={`${styles.btnAction} ${styles.btnClose}`}
                        iconProps={{ iconName: 'Cancel' }}
                        text="Close"
                        onClick={props.onClose}
                    />
                </div>
            </div>

            <div className={styles.tabHeader}>
                <Pivot
                    selectedKey={activeTab}
                    onLinkClick={(item) => setActiveTab(item?.props.itemKey || 'detail')}
                    styles={{
                        root: { borderBottom: '1px solid #eee', background: '#fff', padding: '0 16px' }
                    }}
                >
                    <PivotItem headerText="DETAIL" itemKey="detail" />
                    <PivotItem headerText="ATTACHMENTS" itemKey="attachments" />
                    <PivotItem headerText="ACTIONS" itemKey="actions" />
                </Pivot>
            </div>

            {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>{error}</MessageBar>}

            <div className={styles.scrollContent}>
                {activeTab === 'detail' && (
                    <React.Fragment>
                        <div className={styles.leftColumn}>
                            <Section title="GENERAL INFO" id="general">
                                <ComboBox
                                    label="Type"
                                    options={lookupOptions.trainingTypes.length > 0 ? lookupOptions.trainingTypes : TYPE_OPTIONS}
                                    selectedKey={formData.Type}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => {
                                        handleFieldChange('Type', opt?.key);
                                        if (opt) handleFieldChange('Title', opt.text);
                                    }}
                                    required
                                />
                                <TextField
                                    label="Training For"
                                    value={formData.TrainingFor || ''}
                                    onChange={(_, val) => handleFieldChange('TrainingFor', val)}
                                />
                                <TextField
                                    label="Training Type"
                                    value={formData.TrainingType || ''}
                                    description="Calculated from Type"
                                    readOnly
                                />
                                <ComboBox
                                    label="Participant"
                                    options={lookupOptions.participants}
                                    selectedKey={formData.Participant?.Id}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('Participant', { Id: opt?.key, Title: opt?.text })}
                                    required
                                />
                                <DatePicker label="Schedule Date" value={formData.ScheduledDate ? new Date(formData.ScheduledDate) : undefined} onSelectDate={(date) => handleFieldChange('ScheduledDate', date?.toISOString())} />
                                <ComboBox
                                    label="Status"
                                    options={STATUS_OPTIONS}
                                    selectedKey={formData.Status}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('Status', opt?.key)}
                                />

                                <div style={{ marginTop: 20, display: 'flex', gap: 10, flexDirection: 'column' }}>
                                    <PrimaryButton
                                        iconProps={{ iconName: 'Download' }}
                                        text="DOWNLOAD CERTIFICATE"
                                        style={{ background: '#2B579A', border: 'none' }}
                                        onClick={handleDownloadCertificate}
                                    />
                                    <DefaultButton
                                        iconProps={{ iconName: 'CirclePlus' }}
                                        text="CREATE COMPLIANCE RECORD"
                                        style={{ background: '#000', color: '#fff', border: 'none' }}
                                    />
                                </div>
                            </Section>
                        </div>

                        <div className={styles.rightColumn}>
                            <Section title="OUTCOME" id="outcome">
                                <TextField label="Invitation Status" value={formData.InvitationStatus || ''} onChange={(_, val) => handleFieldChange('InvitationStatus', val)} />
                                <div style={{ padding: '8px 0', borderBottom: '1px solid #eff0f2', marginBottom: 12 }}>
                                    <DefaultButton
                                        iconProps={{ iconName: 'Mail' }}
                                        text="SEND INVITATION"
                                        disabled
                                        styles={{ root: { border: 'none', background: '#e1dfdd', color: '#a19f9d' } }}
                                    />
                                </div>
                                <TextField label="Induction Link" value={formData.InductionLink?.Url || ''} onChange={(_, val) => handleFieldChange('InductionLink', { ...formData.InductionLink, Url: val, Description: val })} />
                                <ComboBox
                                    label="Participation Status"
                                    options={PARTICIPANT_STATUS_OPTIONS}
                                    selectedKey={formData.ParticipantsStatus}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('ParticipantsStatus', opt?.key)}
                                />

                                <div style={{ marginTop: 12 }}>
                                    <PrimaryButton
                                        iconProps={{ iconName: 'Search' }}
                                        text="SHOW RESULT"
                                        style={{ background: '#3b5d82', border: 'none' }}
                                        onClick={handleShowResult}
                                    />
                                </div>

                            </Section>

                            <Section title="INTERNAL" id="internal">
                                <ComboBox
                                    label="Business Profile"
                                    options={lookupOptions.businessProfiles}
                                    selectedKey={formData.BusinessProfile?.Id}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('BusinessProfile', { Id: opt?.key, Title: opt?.text })}
                                />
                                <ComboBox
                                    label="Manager"
                                    options={lookupOptions.employees}
                                    selectedKey={formData.Manager?.Id}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('Manager', { Id: opt?.key, Title: opt?.text })}
                                />
                                <ComboBox
                                    label="Supervisor"
                                    options={lookupOptions.employees}
                                    selectedKey={formData.Supervisors?.Id}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('Supervisors', { Id: opt?.key, Title: opt?.text })}
                                />
                                <ComboBox
                                    label="Coordinator"
                                    options={lookupOptions.employees}
                                    selectedKey={formData.Coordinator?.Id}
                                    allowFreeform={false}
                                    autoComplete='on'
                                    onChange={(_, opt) => handleFieldChange('Coordinator', { Id: opt?.key, Title: opt?.text })}
                                />
                            </Section>

                            <Section title="TIMELINE" id="timeline">
                                {renderTimeline()}
                            </Section>
                        </div>
                    </React.Fragment>
                )}

                {activeTab === 'attachments' && (
                    <div className={styles.attachmentsTab}>
                        {/* ── Toolbar ── */}
                        <div className={styles.attachmentsToolbar}>
                            <span className={styles.documentsTitle}>
                                Documents
                                {libraryDocs.length > 0 && (
                                    <span className={styles.docsBadge}>{libraryDocs.length}</span>
                                )}
                            </span>
                            <div className={styles.attachmentsActions}>
                                <DefaultButton iconProps={{ iconName: 'Upload' }} text="Upload" onClick={() => setIsUploadModalOpen(true)} />
                                <DefaultButton iconProps={{ iconName: 'Delete' }} text="Delete" disabled={selectedDocIds.length === 0 || saving} onClick={deleteSelectedDocs} />
                                <DefaultButton iconProps={{ iconName: 'Export' }} text="Export" menuProps={exportMenuProps} disabled={libraryDocs.length === 0} />
                                <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={fetchLibraryDocs} />
                            </div>
                        </div>

                        {/* ── Selection bar ── */}
                        {selectedDocIds.length > 0 && (
                            <div className={styles.docsSelectionBar}>
                                <Icon iconName="CheckboxComposite" style={{ fontSize: 14, color: '#0078d4' }} />
                                <span>{selectedDocIds.length} document{selectedDocIds.length !== 1 ? 's' : ''} selected</span>
                                <button className={styles.docsClearBtn} onClick={() => setSelectedDocIds([])}>Clear selection</button>
                            </div>
                        )}

                        {/* ── Custom table ── */}
                        <div className={styles.docsTableWrapper}>
                            <table className={styles.docsTable}>
                                <thead>
                                    <tr className={styles.docsHeaderRow}>
                                        <th className={styles.docsColNum}>#</th>
                                        <th className={styles.docsColCheck}>
                                            <input
                                                ref={selectAllRef}
                                                type="checkbox"
                                                checked={isAllSelected}
                                                onChange={handleSelectAll}
                                                title="Select all"
                                            />
                                        </th>
                                        <th>Title</th>
                                        <th>Document Type</th>
                                        <th>File Name</th>
                                        <th>Description</th>
                                        <th>
                                            <span className={styles.thWithIcon}>
                                                Date Uploaded
                                                <Icon iconName="Filter" style={{ fontSize: 11, color: '#888', marginLeft: 4 }} />
                                            </span>
                                        </th>
                                        <th>
                                            <span className={styles.thWithIcon}>
                                                Uploaded By
                                                <Icon iconName="Filter" style={{ fontSize: 11, color: '#888', marginLeft: 4 }} />
                                            </span>
                                        </th>
                                        <th className={styles.docsColAction} />
                                    </tr>
                                    <tr className={styles.docsFilterRow}>
                                        <td /><td />
                                        <td><input className={styles.docsFilterInput} type="text" placeholder="Search..." value={columnFilters.title || ''} onChange={e => setColumnFilters(f => ({ ...f, title: e.target.value }))} /></td>
                                        <td><input className={styles.docsFilterInput} type="text" placeholder="Search..." value={columnFilters.docType || ''} onChange={e => setColumnFilters(f => ({ ...f, docType: e.target.value }))} /></td>
                                        <td><input className={styles.docsFilterInput} type="text" placeholder="Search..." value={columnFilters.fileName || ''} onChange={e => setColumnFilters(f => ({ ...f, fileName: e.target.value }))} /></td>
                                        <td><input className={styles.docsFilterInput} type="text" placeholder="Search..." value={columnFilters.description || ''} onChange={e => setColumnFilters(f => ({ ...f, description: e.target.value }))} /></td>
                                        <td />
                                        <td><input className={styles.docsFilterInput} type="text" placeholder="Search..." value={columnFilters.uploadedBy || ''} onChange={e => setColumnFilters(f => ({ ...f, uploadedBy: e.target.value }))} /></td>
                                        <td />
                                    </tr>
                                </thead>
                                <tbody>
                                    {filteredDocs.length === 0 ? (
                                        <tr>
                                            <td colSpan={9} className={styles.docsNoData}>
                                                {Object.values(columnFilters).some(v => v)
                                                    ? 'No documents match the current filters.'
                                                    : 'No data to display'}
                                            </td>
                                        </tr>
                                    ) : filteredDocs.map((doc, idx) => {
                                        const isSelected = selectedDocIds.includes(doc.Id);
                                        return (
                                            <tr key={doc.Id} className={isSelected ? styles.docsRowSelected : styles.docsRow}>
                                                <td className={styles.docsColNum}>{idx + 1}</td>
                                                <td className={styles.docsColCheck}>
                                                    <input type="checkbox" checked={isSelected} onChange={() => handleSelectOne(doc.Id)} />
                                                </td>
                                                <td>
                                                    <a href={doc.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                                                        {doc.Title || doc.FileName}
                                                    </a>
                                                </td>
                                                <td>{doc.DocumentType?.Name || '—'}</td>
                                                <td className={styles.docsFileName}>{doc.FileName}</td>
                                                <td className={styles.docsDesc}>{doc.Description || '—'}</td>
                                                <td className={styles.docsDate}>
                                                    {new Date(doc.Created).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}
                                                </td>
                                                <td>{doc.Author?.Title || '—'}</td>
                                                <td className={styles.docsColAction}>
                                                    <IconButton
                                                        iconProps={{ iconName: 'Edit' }}
                                                        title="Edit document"
                                                        styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14 } }}
                                                        onClick={() => { setEditingDoc({ ...doc }); setIsEditModalOpen(true); }}
                                                    />
                                                </td>
                                            </tr>
                                        );
                                    })}
                                </tbody>
                            </table>
                        </div>

                        {/* ── Status line ── */}
                        <div className={styles.docsStatusBar}>
                            {filteredDocs.length !== libraryDocs.length
                                ? `Showing ${filteredDocs.length} of ${libraryDocs.length} documents`
                                : `${libraryDocs.length} document${libraryDocs.length !== 1 ? 's' : ''}`}
                        </div>
                    </div>
                )}

                {activeTab === 'actions' && props.item?.Id && (
                    <div className={styles.actionsTab}>
                        <div style={{ padding: '0 10px' }}>
                            <ToDoModuleAny
                                context={props.context}
                                filterRecordId={props.item.Id}
                                filterTrainingInductionValue={actionTrainingInductionValue}
                                defaultRegarding="Training & Inductions"
                                defaultDynamicFieldValue={actionTrainingInductionValue}
                                isSubGrid={true}
                            />
                        </div>
                    </div>
                )}
            </div>
            {isUploadModalOpen && (
                <Modal
                    isOpen={isUploadModalOpen}
                    onDismiss={() => { setIsUploadModalOpen(false); setUploadQueue([]); setIsDragging(false); }}
                    isBlocking={false}
                    containerClassName={styles.uploadModal}
                >
                    {/* ── Header ── */}
                    <div className={styles.uploadHeader}>
                        <div className={styles.uploadHeaderLeft}>
                            <Icon iconName="CloudUpload" style={{ fontSize: 22, color: '#fff', opacity: 0.95 }} />
                            <div>
                                <div className={styles.uploadHeaderTitle}>Upload Documents</div>
                                <div className={styles.uploadHeaderSub}>
                                    Record #{props.item?.Id}{formData.Title ? ` · ${formData.Title}` : ''}
                                </div>
                            </div>
                        </div>
                        <IconButton
                            iconProps={{ iconName: 'Cancel' }}
                            styles={{ root: { color: 'rgba(255,255,255,0.8)' }, icon: { color: 'rgba(255,255,255,0.9)' }, rootHovered: { background: 'rgba(255,255,255,0.15)', color: '#fff' } }}
                            onClick={() => { setIsUploadModalOpen(false); setUploadQueue([]); }}
                        />
                    </div>

                    {/* ── Body ── */}
                    <div className={styles.uploadBody}>
                        {/* Drop zone */}
                        <div
                            className={`${styles.dropZone}${isDragging ? ` ${styles.dropZoneActive}` : ''}`}
                            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                            onDragLeave={() => setIsDragging(false)}
                            onDrop={(e) => {
                                e.preventDefault();
                                setIsDragging(false);
                                const dropped = Array.from(e.dataTransfer.files);
                                setUploadQueue(prev => [...prev, ...dropped.map(f => ({ file: f, title: f.name.split('.').slice(0, -1).join('.'), docType: '', description: '' }))]);
                            }}
                        >
                            <Icon iconName={isDragging ? 'CloudAdd' : 'CloudUpload'} className={styles.dropZoneIcon} />
                            <p className={styles.dropZoneTitle}>{isDragging ? 'Drop files here' : 'Drag & drop files here'}</p>
                            <p className={styles.dropZoneSub}>or</p>
                            <PrimaryButton
                                iconProps={{ iconName: 'FolderOpen' }}
                                text="Browse Files"
                                onClick={() => document.getElementById('fileInput')?.click()}
                            />
                            <input type="file" multiple id="fileInput" style={{ display: 'none' }} onChange={handleFileUpload} />
                            <p className={styles.dropZoneHint}>PDF, Word, Excel, images and more are supported</p>
                        </div>

                        {/* File queue */}
                        {uploadQueue.length > 0 && (
                            <div className={styles.fileQueue}>
                                <div className={styles.fileQueueTitle}>
                                    <Icon iconName="Attach" style={{ fontSize: 13, color: '#0078d4', marginRight: 6 }} />
                                    Files to upload ({uploadQueue.length})
                                </div>
                                {uploadQueue.map((item, idx) => (
                                    <div key={idx} className={styles.fileQueueItem}>
                                        <Icon iconName={getFileIcon(item.file.name)} className={styles.fileQueueIcon} />
                                        <div className={styles.fileQueueItemInfo}>
                                            <span className={styles.fileQueueItemName}>{item.file.name}</span>
                                            <span className={styles.fileQueueItemSize}>{formatFileSize(item.file.size)}</span>
                                        </div>
                                        <IconButton
                                            iconProps={{ iconName: 'Delete' }}
                                            title="Remove file"
                                            styles={{ root: { height: 30, width: 30 }, icon: { fontSize: 14, color: '#a4262c' }, rootHovered: { background: '#fde7e9' } }}
                                            onClick={() => setUploadQueue(prev => prev.filter((_, i) => i !== idx))}
                                        />
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>

                    {/* ── Footer ── */}
                    <div className={styles.modalFooter}>
                        <DefaultButton text="Cancel" onClick={() => { setIsUploadModalOpen(false); setUploadQueue([]); }} />
                        <PrimaryButton
                            iconProps={{ iconName: saving ? 'ProgressRingDots' : 'CloudUpload' }}
                            text={saving ? 'Uploading…' : `Upload ${uploadQueue.length} file${uploadQueue.length !== 1 ? 's' : ''}`}
                            disabled={saving || uploadQueue.length === 0}
                            onClick={saveUploadedDocs}
                        />
                    </div>
                </Modal>
            )}

            {isEditModalOpen && editingDoc && (
                <Modal
                    isOpen={isEditModalOpen}
                    onDismiss={() => setIsEditModalOpen(false)}
                    containerClassName={styles.uploadModal}
                >
                    {/* ── Header ── */}
                    <div className={styles.uploadHeader}>
                        <div className={styles.uploadHeaderLeft}>
                            <Icon iconName={getFileIcon(editingDoc.FileName)} style={{ fontSize: 22, color: '#fff', opacity: 0.95 }} />
                            <div className={styles.uploadHeaderTitle}>Edit Document</div>
                        </div>
                        <IconButton
                            iconProps={{ iconName: 'Cancel' }}
                            styles={{ root: { color: 'rgba(255,255,255,0.8)' }, icon: { color: 'rgba(255,255,255,0.9)' }, rootHovered: { background: 'rgba(255,255,255,0.15)', color: '#fff' } }}
                            onClick={() => setIsEditModalOpen(false)}
                        />
                    </div>

                    {/* ── Body ── */}
                    <div className={styles.uploadBody}>
                        {/* Document info card */}
                        <div className={styles.editDocInfo}>
                            <Icon iconName={getFileIcon(editingDoc.FileName)} style={{ fontSize: 30, color: '#0078d4', flexShrink: 0 }} />
                            <div style={{ flex: 1, minWidth: 0 }}>
                                <div className={styles.editDocFileName}>{editingDoc.FileName}</div>
                                <div className={styles.editDocMeta}>
                                    Uploaded by {editingDoc.Author?.Title || 'Unknown'} &nbsp;·&nbsp;
                                    {new Date(editingDoc.Created).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}
                                </div>
                            </div>
                            <a
                                href={editingDoc.ServerRelativeUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className={styles.editDocLink}
                            >
                                <Icon iconName="OpenInNewTab" style={{ fontSize: 11, marginRight: 4 }} />
                                View file
                            </a>
                        </div>

                        <TextField
                            label="Title"
                            value={editingDoc.Title}
                            onChange={(_, v) => setEditingDoc({ ...editingDoc, Title: v || '' })}
                            required
                            styles={{ root: { marginTop: 4 } }}
                        />
                        <ComboBox
                            label="Document Type"
                            options={lookupOptions.documentTypes}
                            selectedKey={editingDoc.DocumentType?.Id || null}
                            allowFreeform={false}
                            autoComplete="on"
                            onChange={(_, opt) => setEditingDoc({ ...editingDoc, DocumentType: { Id: opt?.key as number, Name: opt?.text || '' } })}
                            styles={{ root: { marginTop: 12 } }}
                        />
                        <TextField
                            label="Description"
                            multiline
                            rows={4}
                            value={editingDoc.Description}
                            onChange={(_, v) => setEditingDoc({ ...editingDoc, Description: v || '' })}
                            styles={{ root: { marginTop: 12 } }}
                        />
                    </div>

                    {/* ── Footer ── */}
                    <div className={styles.modalFooter}>
                        <DefaultButton text="Cancel" onClick={() => setIsEditModalOpen(false)} />
                        <PrimaryButton
                            iconProps={{ iconName: saving ? 'ProgressRingDots' : 'Save' }}
                            text={saving ? 'Updating…' : 'Update Document'}
                            disabled={saving}
                            onClick={handleUpdateMetadata}
                        />
                    </div>
                </Modal>
            )}
        </div>
    );
};

export default TrainingInductionForm;

