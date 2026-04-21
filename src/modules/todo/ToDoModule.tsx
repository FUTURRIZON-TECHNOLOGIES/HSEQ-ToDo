import * as React from 'react';
import { IToDoItem } from '../../webparts/swf/models/IToDoItem';
import { SPService } from '../../webparts/swf/services/SPService';
import GenericGrid from '../../webparts/swf/components/Shared/GenericGrid';
import { IColumn, Panel, PanelType, Icon } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ToDoForm from './ToDoForm';

export interface IToDoModuleProps {
    context: WebPartContext;
}

// ── Comprehensive Export Definitions ──────────────────────────────────────────
const formatDate = (d: string | undefined) => {
    if (!d) return '';
    const date = new Date(d);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
};

const stripHtml = (html: string | undefined): string => {
    if (!html) return '';
    return html.replace(/(<([^>]+)>)/gi, "").trim();
};

const PAGE_SIZE = 100;

const ToDoModule: React.FC<IToDoModuleProps> = ({ context }) => {
    const [items, setItems] = React.useState<IToDoItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [selectedItem, setSelectedItem] = React.useState<IToDoItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);
    const [formVersion, setFormVersion] = React.useState(0);

    // ── Pagination & Search State ─────────────────────────────────────────────
    const [currentPage, setCurrentPage] = React.useState(1);
    const [totalCount, setTotalCount] = React.useState(0);
    const [searchQuery, setSearchQuery] = React.useState('');
    const [sortConfig, setSortConfig] = React.useState<{ field: string; isAscending: boolean }>({
        field: 'Id',
        isAscending: true
    });
    const totalPages = Math.max(1, Math.ceil(totalCount / PAGE_SIZE));

    const spService = React.useMemo(() => new SPService(context), [context]);

    const fetchData = React.useCallback(async (page: number, search: string, sortField: string, isAsc: boolean) => {
        setLoading(true);
        try {
            const [data, total] = await Promise.all([
                spService.getToDoItemsPaged(page, PAGE_SIZE, search, sortField, isAsc),
                spService.getToDoTotalCount(search)
            ]);
            const mappedItems = data.map(item => ({
                ...item,
                StatusValue: item.Status?.Title || item.Status?.Name || item.StatusId?.toString() || item.Status?.toString() || "",
                CategoryValue: item.Category?.Title || item.Category?.Name || item.CategoryId?.toString() || item.Category?.toString() || "",
                ClassificationValue: item.Classification?.Title || item.Classification?.Name || item.ClassificationId?.toString() || item.Classification?.toString() || "",
                PriorityValue: item.Priority?.Title || item.Priority?.Name || item.PriorityId?.toString() || item.Priority?.toString() || ""
            }));
            setItems(mappedItems);
            setTotalCount(total);
        } catch (e) {
            console.error('[ToDoModule] Fetch failed', e);
        } finally {
            setLoading(false);
        }
    }, [spService]);

    React.useEffect(() => {
        fetchData(currentPage, searchQuery, sortConfig.field, sortConfig.isAscending);
    }, [currentPage, searchQuery, sortConfig]);

    // Keep a stable "refresh current page" reference for CRUD callbacks
    const fetchItems = React.useCallback(() => 
        fetchData(currentPage, searchQuery, sortConfig.field, sortConfig.isAscending), 
    [fetchData, currentPage, searchQuery, sortConfig]);

    // ── Grid Column Definitions ───────────────────────────────────────────────
    const columns: IColumn[] = [
        {
            key: 'Id', name: 'ID', fieldName: 'Id',
            minWidth: 40, maxWidth: 55, isResizable: true
        },
        {
            key: 'Regarding', name: 'Regarding', fieldName: 'Regarding',
            minWidth: 120, maxWidth: 180, isResizable: true,
            onRender: (item: IToDoItem) => {
                if (typeof item.Regarding === 'object' && item.Regarding !== null) {
                    return <span>{(item.Regarding as any).Title || '—'}</span>;
                }
                return <span>{item.Regarding || '—'}</span>;
            }
        },
        {
            key: 'Status', name: 'Status', fieldName: 'StatusValue',
            minWidth: 90, maxWidth: 120, isResizable: true,
            onRender: (item: any) => <span>{item.StatusValue || '—'}</span>
        },
        {
            key: 'Category', name: 'Category', fieldName: 'CategoryValue',
            minWidth: 100, maxWidth: 130, isResizable: true,
            onRender: (item: any) => <span>{item.CategoryValue || '—'}</span>
        },
        {
            key: 'Classification', name: 'Classification', fieldName: 'ClassificationValue',
            minWidth: 100, maxWidth: 130, isResizable: true,
            onRender: (item: any) => <span>{item.ClassificationValue || '—'}</span>
        },
        {
            key: 'Priority', name: 'Priority', fieldName: 'PriorityValue',
            minWidth: 80, maxWidth: 100, isResizable: true,
            onRender: (item: any) => <span>{item.PriorityValue || '—'}</span>
        },
        {
            key: 'TaskOwner', name: 'Task Owner', fieldName: 'TaskOwner',
            minWidth: 120, maxWidth: 160, isResizable: true,
            onRender: (item: IToDoItem) => {
                const title = item.TaskOwner?.Title || '';
                if (!title) return <span>—</span>;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Contact" style={{ fontSize: 12, color: '#0078d4' }} />
                        <span>{title}</span>
                    </div>
                );
            }
        },
        {
            key: 'AssigneeInternal', name: 'Assignee Internal', fieldName: 'AssigneeInternal',
            minWidth: 120, maxWidth: 160, isResizable: true,
            onRender: (item: IToDoItem) => {
                const title = item.AssigneeInternal?.Title || '';
                if (!title) return <span>—</span>;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Contact" style={{ fontSize: 12, color: '#0078d4' }} />
                        <span>{title}</span>
                    </div>
                );
            }
        },
        {
            key: 'DueDate', name: 'Due Date', fieldName: 'DueDate',
            minWidth: 90, maxWidth: 110, isResizable: true,
            onRender: (item: IToDoItem) => <span>{formatDate(item.DueDate)}</span>
        },
        {
            key: 'StartDate', name: 'Start Date', fieldName: 'StartDate',
            minWidth: 90, maxWidth: 110, isResizable: true,
            onRender: (item: IToDoItem) => <span>{formatDate(item.StartDate)}</span>
        },
        {
            key: 'CompletionDate', name: 'Completion Date', fieldName: 'CompletionDate',
            minWidth: 110, maxWidth: 130, isResizable: true,
            onRender: (item: IToDoItem) => <span>{formatDate(item.CompletionDate)}</span>
        },
        {
            key: 'CompletedPercent', name: '% Done', fieldName: 'CompletedPercent',
            minWidth: 65, maxWidth: 80, isResizable: true,
            onRender: (item: IToDoItem) => {
                const pct = item.CompletedPercent ?? 0;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                        <div style={{ flex: 1, height: 6, background: '#e0e0e0', borderRadius: 3, minWidth: 36 }}>
                            <div style={{
                                width: `${Math.min(pct, 100)}%`, height: '100%',
                                background: pct >= 100 ? '#107C10' : pct >= 50 ? '#0078D4' : '#F7901E',
                                borderRadius: 3, transition: 'width 0.3s'
                            }} />
                        </div>
                        <span style={{ fontSize: 11, color: '#555', whiteSpace: 'nowrap' }}>{pct}%</span>
                    </div>
                );
            }
        },
        {
            key: 'Attachments', name: 'Attachments', fieldName: 'AttachmentFiles',
            minWidth: 120, maxWidth: 220, isResizable: true,
            onRender: (item: IToDoItem) => {
                if (!item.AttachmentFiles || item.AttachmentFiles.length === 0) return <span>—</span>;
                return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                        {item.AttachmentFiles.map((file, idx) => (
                            <a key={idx} href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" style={{ textDecoration: 'none', color: '#0078D4', display: 'flex', alignItems: 'center', gap: 4 }}>
                                <Icon iconName="Attach" style={{ fontSize: 12 }} />
                                <span style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }} title={file.FileName}>{file.FileName}</span>
                            </a>
                        ))}
                    </div>
                );
            }
        }
    ];

    // ── Flatten items for export (FULL DATA) ──────────────────────────────────
    const flattenItems = (rawItems: any[]) => rawItems.map(item => ({
        'ID': item.Id,
        'Subject': item.Title || '',
        'Regarding': item.Regarding || '',
        'Status': item.Status?.Title || '',
        'Category': item.Category?.Title || '',
        'Classification': item.Classification?.Title || '',
        'Priority': item.Priority?.Title || '',
        'Task Owner': item.TaskOwner?.Title || '',
        'Assignee Internal': item.AssigneeInternal?.Title || '',
        'Assignee External': item.AssigneeExternal?.Title || '',
        'Due Date': formatDate(item.DueDate),
        'Start Date': formatDate(item.StartDate),
        'Completion Date': formatDate(item.CompletionDate),
        '% Done': `${item.CompletedPercent ?? 0}%`,
        'Resolution': stripHtml(item.Resolution),
        'Description': stripHtml(item.Description),
        'Email Notifications': item.EmailNotifications ? 'Yes' : 'No',
        'Created By': item.Author?.Title || '',
        'Created On': formatDate(item.Created),
        'Modified By': item.Editor?.Title || '',
        'Modified On': formatDate(item.Modified)
    }));

    // ── Export Handlers (FETCH FULL FILTERED DATA) ───────────────────────────
    const handleExportExcel = async (pagedData: any[]) => {
        setLoading(true);
        try {
            const allData = await spService.getToDoItemsFiltered(searchQuery, sortConfig.field, sortConfig.isAscending);
            const { ExportService } = await import('../../webparts/swf/services/ExportService');
            ExportService.exportToExcel(flattenItems(allData), 'ASP_Assist_Group_Tasks');
        } finally { setLoading(false); }
    };

    const handleExportCSV = async (pagedData: any[]) => {
        setLoading(true);
        try {
            const allData = await spService.getToDoItemsFiltered(searchQuery, sortConfig.field, sortConfig.isAscending);
            const { ExportService } = await import('../../webparts/swf/services/ExportService');
            ExportService.exportToCSV(flattenItems(allData), 'ASP_Assist_Group_Tasks');
        } finally { setLoading(false); }
    };

    const handleExportPDF = async (pagedData: any[]) => {
        setLoading(true);
        try {
            const allData = await spService.getToDoItemsFiltered(searchQuery, sortConfig.field, sortConfig.isAscending);
            const { ExportService } = await import('../../webparts/swf/services/ExportService');
            const flattened = flattenItems(allData);
            const headers = Object.keys(flattened[0] || {});
            ExportService.exportToPDF(flattened, 'ASP Assist Group Tasks', headers, headers);
        } finally { setLoading(false); }
    };

    const handleExportZip = async (pagedData: any[]) => {
        setLoading(true);
        try {
            const allData = await spService.getToDoItemsFiltered(searchQuery, sortConfig.field, sortConfig.isAscending);
            const { ExportService } = await import('../../webparts/swf/services/ExportService');
            const flattened = flattenItems(allData);
            const headers = Object.keys(flattened[0] || {});
            await ExportService.exportToZip(flattened, 'ASP_Assist_Group_Tasks', headers, headers);
        } catch (e) {
            console.error('ZIP export failed', e);
        } finally { setLoading(false); }
    };

    // ── CRUD Handlers ─────────────────────────────────────────────────────────
    const handleDelete = async (selectedItems: any[]) => {
        if (confirm(`Delete ${selectedItems.length} item(s)?`)) {
            for (const item of selectedItems) await spService.deleteToDoItem(item.Id);
            await fetchItems();
        }
    };

    const handleNew = () => { 
        setSelectedItem(undefined); 
        setFormVersion(prev => prev + 1);
        setIsPanelOpen(true); 
    };
    const handleEdit = (item: IToDoItem) => { setSelectedItem(item); setIsPanelOpen(true); };

    const handleSave = async (payload: any, mode: 'stay' | 'close' | 'new') => {
        try {
            // Track the ID from the current selection if we are in edit mode
            let resultItemId = selectedItem?.Id;
            const isUpdate = !!resultItemId;

            if (isUpdate) {
                // Perform Update
                await spService.updateToDoItem(resultItemId!, payload);
            } else {
                // Perform Add
                const result = await spService.addToDoItem(payload);
                resultItemId = result.data?.Id || result.Id;

                // Handle pending attachments for new items
                if (payload.PendingAttachments && payload.PendingAttachments.length > 0 && resultItemId) {
                    for (const file of payload.PendingAttachments) {
                        try {
                            await spService.uploadAttachment(resultItemId, file);
                        } catch (err) {
                            console.error('Failed to upload pending attachment', err);
                        }
                    }
                }
            }
            
            // Refresh counts and current page data
            const newTotalCount = await spService.getToDoTotalCount(searchQuery);
            setTotalCount(newTotalCount);

            let targetPage = currentPage;
            const maxPage = Math.max(1, Math.ceil(newTotalCount / PAGE_SIZE));
            if (targetPage > maxPage) targetPage = maxPage;

            const refreshedData = await spService.getToDoItemsPaged(
                targetPage, 
                PAGE_SIZE, 
                searchQuery, 
                sortConfig.field, 
                sortConfig.isAscending
            );
            setItems(refreshedData);

            // Handle panel state and selection based on the save mode
            if (mode === 'close') {
                setIsPanelOpen(false);
                setSelectedItem(undefined);
            } else if (mode === 'new') {
                setSelectedItem(undefined);
                setFormVersion(prev => prev + 1);
            } else if (mode === 'stay' && resultItemId) {
                // Attempt to find the fresh item to keep the form updated
                const freshItem = refreshedData.find(i => i.Id === resultItemId);
                if (freshItem) {
                    setSelectedItem(freshItem);
                } else {
                    // For both Add and Update, if freshItem is not found (search lag),
                    // ensure we track the ID so the form transitions to Edit mode.
                    // If isUpdate was false, we use the new ID from SharePoint result.
                    setSelectedItem(prev => ({ 
                        ...((prev || {}) as any), 
                        ...payload, 
                        Id: resultItemId 
                    } as IToDoItem));
                }
            }
        } catch (e) {
            console.error('Save failed', e);
            alert('Save failed. Check browser console for details.');
        }
    };
    
    const handleRefresh = async () => {
        if (selectedItem?.Id) {
            try {
                // Fetch fresh copy from SharePoint
                const items = await spService.getToDoItems();
                const fresh = items.find(i => i.Id === selectedItem.Id);
                if (fresh) setSelectedItem(fresh);
            } catch (e) {
                console.error('Refresh failed', e);
            }
        } else {
            // On a new form, refresh simply clears the form
            setFormVersion(prev => prev + 1);
        }
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
                    setCurrentPage(1); // Reset to first page on search
                }}
                onExportExcel={handleExportExcel}
                onExportCSV={handleExportCSV}
                onExportPDF={handleExportPDF}
                onExportZip={handleExportZip}
                currentPage={currentPage}
                totalPages={totalPages}
                totalCount={totalCount}
                pageSize={PAGE_SIZE}
                clientSidePagination={false}
                onPageChange={(page) => setCurrentPage(page)}
                sortField={sortConfig.field}
                isAscending={sortConfig.isAscending}
                onSort={(field, isAsc) => {
                    setSortConfig({ field, isAscending: isAsc });
                    setCurrentPage(1); // Reset to first page on sort change
                }}
            />

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                type={PanelType.custom}
                customWidth="1100px"
                onRenderHeader={() => null}
                isLightDismiss={false}
                styles={{
                    content: { padding: 0 },
                    scrollableContent: { overflow: 'hidden' },
                    commands: { display: 'none' }
                }}
            >
                <ToDoForm
                    key={`${selectedItem?.Id || 'new'}-${formVersion}`}
                    item={selectedItem}
                    spService={spService}
                    context={context}
                    onSave={handleSave}
                    onRefresh={handleRefresh}
                    onClose={() => setIsPanelOpen(false)}
                />
            </Panel>
        </React.Fragment>
    );
};

export default ToDoModule;
