import * as React from 'react';
import { ITrainingInductionItem } from '../../../models/ITrainingInductionItem';
import { SPService } from '../../../services/SPService';
import GenericGrid from '../../Shared/GenericGrid';
import { IColumn, Panel, PanelType, Icon, IDropdownOption } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import TrainingInductionForm from './TrainingInductionForm';
import TrainingInductionNewForm from './TrainingInductionNewForm';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
const JSZip: any = require('jszip');

export interface ITrainingInductionModuleProps {
    context: WebPartContext;
}

const formatDate = (d: string | undefined): string => {
    if (!d) return '';
    const date = new Date(d);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
};

const PAGE_SIZE = 100;

const TrainingInductionModule: React.FC<ITrainingInductionModuleProps> = ({ context }) => {
    const [items, setItems] = React.useState<ITrainingInductionItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [selectedItem, setSelectedItem] = React.useState<ITrainingInductionItem | null>(null);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [currentSaveId, setCurrentSaveId] = React.useState<number | null>(null);

    const [panelTab, setPanelTab] = React.useState<'detail' | 'attachments'>('detail');
    const [currentPage, setCurrentPage] = React.useState(1);
    const [totalCount, setTotalCount] = React.useState(0);
    const [searchQuery, setSearchQuery] = React.useState('');
    const [sortConfig, setSortConfig] = React.useState<{ field: string; isAscending: boolean }>({
        field: 'Id',
        isAscending: true
    });
    const [lookupOptions, setLookupOptions] = React.useState<{
        participants: IDropdownOption[],
        businessProfiles: IDropdownOption[],
        employees: IDropdownOption[],
        trainingTypes: IDropdownOption[]
    }>({
        participants: [],
        businessProfiles: [],
        employees: [],
        trainingTypes: []
    });
    const totalPages = Math.max(1, Math.ceil(totalCount / PAGE_SIZE));

    const spService = React.useMemo(() => new SPService(context), [context]);

    const fetchData = React.useCallback(async (page: number, search: string, sortField: string, isAsc: boolean): Promise<void> => {
        setLoading(true);
        try {
            const [data, total] = await Promise.all([
                spService.getTrainingInductionItemsPaged(page, PAGE_SIZE, search, sortField, isAsc),
                spService.getTrainingInductionTotalCount(search)
            ]);
            setItems(data);
            setTotalCount(total);
        } catch (e) {
            console.error('[TrainingInductionModule] Fetch failed', e);
        } finally {
            setLoading(false);
        }
    }, [spService]);

    const loadLookups = React.useCallback(async (): Promise<void> => {
        try {
            const [pts, bps, emps, ttypes] = await Promise.all([
                spService.getLookupOptions('Contacts', 'Employee Name'),
                spService.getLookupOptions('BusinessProfiles', 'BusinessProfile'),
                spService.getLookupOptions('Employees', 'Employee Name'),
                spService.getLookupOptions('TrainingTypes', 'Title')
            ]);
            setLookupOptions({
                participants: pts.map(p => ({ key: p.Id, text: p.Title })),
                businessProfiles: bps.map(b => ({ key: b.Id, text: b.Title })),
                employees: emps.map(e => ({ key: e.Id, text: e.Title })),
                trainingTypes: ttypes.map(t => ({ key: t.Title, text: t.Title }))
            });
        } catch (e) {
            console.error('Lookup load failed', e);
        }
    }, [spService]);

    React.useEffect(() => {
        fetchData(currentPage, searchQuery, sortConfig.field, sortConfig.isAscending);
        loadLookups();
    }, [currentPage, searchQuery, sortConfig, loadLookups, fetchData]);

    const fetchItems = React.useCallback((): Promise<void> =>
        fetchData(currentPage, searchQuery, sortConfig.field, sortConfig.isAscending),
        [fetchData, currentPage, searchQuery, sortConfig]);

    const columns: IColumn[] = [
        {
            key: 'Id', name: 'ID', fieldName: 'Id',
            minWidth: 40, maxWidth: 60, isResizable: true
        },
        {
            key: 'Type', name: 'Type', fieldName: 'Type',
            minWidth: 150, maxWidth: 250, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{typeof item.Type === 'object' ? (item.Type as any)?.Title || JSON.stringify(item.Type) : (item.Type || '—')}</span>
        },
        {
            key: 'Participant', name: 'Participant', fieldName: 'Participant',
            minWidth: 120, maxWidth: 180, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => {
                const title = item.Participant?.Title || '';
                if (!title) return <span>—</span>;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Contact" style={{ fontSize: 12, color: '#0078d4' }} />
                        <span style={{ color: '#0078d4', cursor: 'pointer', textDecoration: 'underline' }}>{title}</span>
                    </div>
                );
            }
        },
        {
            key: 'ParticipantsStatus', name: 'Participant Status', fieldName: 'ParticipantsStatus',
            minWidth: 100, maxWidth: 130, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{typeof item.ParticipantsStatus === 'object' ? (item.ParticipantsStatus as any)?.Title || JSON.stringify(item.ParticipantsStatus) : (item.ParticipantsStatus || '—')}</span>
        },
        {
            key: 'Company', name: 'Company', fieldName: 'Company',
            minWidth: 120, maxWidth: 200, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{item.Company?.Title || '—'}</span>
        },
        {
            key: 'TrainingFor', name: 'Induction For', fieldName: 'TrainingFor',
            minWidth: 100, maxWidth: 150, isResizable: true
        },
        {
            key: 'Project', name: 'Project', fieldName: 'Project',
            minWidth: 100, maxWidth: 150, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{item.Project?.Title || '—'}</span>
        },
        {
            key: 'ScheduledDate', name: 'Schedule Date', fieldName: 'ScheduledDate',
            minWidth: 100, maxWidth: 120, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{formatDate(item.ScheduledDate)}</span>
        },
        {
            key: 'CompletionDate', name: 'Completion Date', fieldName: 'CompletionDate',
            minWidth: 100, maxWidth: 120, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{formatDate(item.CompletionDate)}</span>
        },
        {
            key: 'Status', name: 'Status', fieldName: 'Status',
            minWidth: 90, maxWidth: 120, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => (
                <span style={{
                    color: item.Status === 'Completed' ? '#107c10' : item.Status === 'In Progress' ? '#0078d4' : '#666',
                    fontWeight: 600
                }}>
                    {typeof item.Status === 'object' ? (item.Status as any)?.Title || 'Scheduled' : (item.Status || 'Scheduled')}
                </span>
            )
        },
        {
            key: 'BusinessProfile', name: 'Business Profile', fieldName: 'BusinessProfile',
            minWidth: 120, maxWidth: 180, isResizable: true,
            onRender: (item: ITrainingInductionItem): JSX.Element => <span>{item.BusinessProfile?.Title || '—'}</span>
        },
        {
            key: 'Attachments', name: '', fieldName: 'Attachments',
            minWidth: 40, maxWidth: 40, isResizable: false,
            onRender: (item: ITrainingInductionItem): JSX.Element => (
                <Icon 
                    iconName="Attach" 
                    style={{ fontSize: 16, color: '#0078d4', cursor: 'pointer' }} 
                    onClick={() => handleViewAttachments(item)}
                    title="View Attachments"
                />
            )
        }
    ];

    const handleDelete = async (selectedItems: any[]) => {
        if (confirm(`Delete ${selectedItems.length} item(s)?`)) {
            for (const item of selectedItems) await spService.deleteTrainingInductionItem(item.Id);
            await fetchItems();
        }
    };

    const handleNew = () => {
        setSelectedItem(null);
        setCurrentSaveId(null);
        setIsPanelOpen(true);
    };
    const handleEdit = (item: ITrainingInductionItem) => {
        setSelectedItem(item);
        setCurrentSaveId(item.Id || null);
        setPanelTab('detail');
        setIsPanelOpen(true);
    };
    const handleViewAttachments = (item: ITrainingInductionItem) => {
        setSelectedItem(item);
        setCurrentSaveId(item.Id || null);
        setPanelTab('attachments');
        setIsPanelOpen(true);
    };

    const handleSave = async (payload: any, mode: 'stay' | 'close' | 'new') => {
        try {
            if (selectedItem || currentSaveId) {
                await spService.updateTrainingInductionItem(selectedItem?.Id || currentSaveId!, payload);
            } else {
                const result = await spService.addTrainingInductionItem(payload);
                if (mode === 'stay' && result) {
                    // Update currentSaveId but DO NOT switch selectedItem
                    // This stays in the "New" form layout but enables "Update" for future saves
                    setCurrentSaveId(result.data?.Id || result.Id);
                }
            }

            await fetchItems();

            if (mode === 'close') {
                setIsPanelOpen(false);
                setSelectedItem(null);
                setCurrentSaveId(null);
            } else if (mode === 'new') {
                setSelectedItem(null);
                setCurrentSaveId(null);
            }
        } catch (e) {
            console.error('Save failed', e);
            throw e;
        }
    };

    // ─── Export Logic ───

    const prepareExportData = async (targetItems: any[], isSelection?: boolean) => {
        let finalItems = targetItems;
        
        // If no records selected, fetch all from SharePoint to bypass threshold
        if (!isSelection) {
            setLoading(true);
            try {
                console.log("[Export] Fetching all records from SharePoint...");
                finalItems = await spService.getAllTrainingInductionItems(searchQuery);
            } catch (e) {
                console.error("[Export] Fetch failed", e);
            } finally {
                setLoading(false);
            }
        }
        
        if (!finalItems || finalItems.length === 0) return [];
        
        return finalItems.map(item => ({
            ID: item.Id,
            Type: typeof item.Type === 'object' ? item.Type?.Title : (item.Type || ''),
            Participant: item.Participant?.Title || '',
            'Participant Status': typeof item.ParticipantsStatus === 'object' ? item.ParticipantsStatus?.Title : (item.ParticipantsStatus || ''),
            Company: item.Company?.Title || '',
            'Induction For': item.TrainingFor || '',
            Project: item.Project?.Title || '',
            'Scheduled Date': formatDate(item.ScheduledDate),
            'Completion Date': formatDate(item.CompletionDate),
            Status: typeof item.Status === 'object' ? item.Status?.Title : (item.Status || ''),
            'Business Profile': item.BusinessProfile?.Title || ''
        }));
    };

    const downloadFile = (blob: Blob, fileName: string) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    };

    const handleExportExcel = async (targetItems: any[], isSelection?: boolean) => {
        const data = await prepareExportData(targetItems, isSelection);
        if (data.length === 0) return;
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");
        const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
        downloadFile(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `training-inductions.xlsx`);
    };

    const handleExportCSV = async (targetItems: any[], isSelection?: boolean) => {
        const data = await prepareExportData(targetItems, isSelection);
        if (data.length === 0) return;
        const ws = XLSX.utils.json_to_sheet(data);
        const csv = XLSX.utils.sheet_to_csv(ws);
        downloadFile(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), `training-inductions.csv`);
    };

    const handleExportPDF = async (targetItems: any[], isSelection?: boolean) => {
        const data = await prepareExportData(targetItems, isSelection);
        if (data.length === 0) return;
        const doc = new jsPDF('l', 'pt', 'a4');
        autoTable(doc, {
            head: [Object.keys(data[0])],
            body: data.map(item => Object.values(item)),
            startY: 40,
            styles: { fontSize: 7, cellPadding: 2 },
            theme: 'grid',
            headStyles: { fillColor: [0, 120, 212] }
        });
        doc.save(`training-inductions.pdf`);
    };

    const handleExportZip = async (targetItems: any[], isSelection?: boolean) => {
        const data = await prepareExportData(targetItems, isSelection);
        if (data.length === 0) return;
        const zip = new JSZip();

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");
        zip.file("training-inductions.xlsx", XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
        zip.file("training-inductions.csv", XLSX.utils.sheet_to_csv(ws));

        const doc = new jsPDF('l', 'pt', 'a4');
        autoTable(doc, {
            head: [Object.keys(data[0])],
            body: data.map(item => Object.values(item)),
            styles: { fontSize: 7 }
        });
        zip.file("training-inductions.pdf", doc.output('blob'));

        const content = await zip.generateAsync({ type: "blob" });
        downloadFile(content, `training-inductions.zip`);
    };

    return (
        <React.Fragment>
            <GenericGrid
                items={items}
                columns={columns}
                loading={loading}
                onNew={handleNew}
                onEdit={handleEdit}
                onDelete={handleDelete}
                onRefresh={fetchItems}
                onSearch={(term) => {
                    setSearchQuery(term);
                    setCurrentPage(1);
                }}
                currentPage={currentPage}
                totalPages={totalPages}
                totalCount={totalCount}
                pageSize={PAGE_SIZE}
                onPageChange={(page) => setCurrentPage(page)}
                sortField={sortConfig.field}
                isAscending={sortConfig.isAscending}
                onSort={(field, isAsc) => {
                    setSortConfig({ field, isAscending: isAsc });
                    setCurrentPage(1);
                }}
                onExportExcel={handleExportExcel}
                onExportCSV={handleExportCSV}
                onExportPDF={handleExportPDF}
                onExportZip={handleExportZip}
            />

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => { setIsPanelOpen(false); setPanelTab('detail'); }}
                type={PanelType.custom}
                customWidth={selectedItem ? "1100px" : "600px"}
                onRenderHeader={() => null}
                isLightDismiss={false}
                styles={{
                    content: { padding: 0 },
                    scrollableContent: { overflow: 'hidden' },
                    commands: { display: 'none' }
                }}
            >
                {selectedItem ? (
                    <TrainingInductionForm
                        item={selectedItem}
                        spService={spService}
                        context={context}
                        onClose={() => setIsPanelOpen(false)}
                        onSave={handleSave}
                        preLoadedLookups={lookupOptions}
                        defaultTab={panelTab}
                    />
                ) : (
                    <TrainingInductionNewForm
                        spService={spService}
                        context={context}
                        onClose={() => setIsPanelOpen(false)}
                        onSave={handleSave}
                        preLoadedLookups={lookupOptions}
                    />
                )}
            </Panel>
        </React.Fragment>
    );
};

export default TrainingInductionModule;
