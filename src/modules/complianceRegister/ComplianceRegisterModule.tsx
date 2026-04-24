import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { IComplianceItem, ILookupSets, IAttachment } from './models/IComplianceItem';
import { ComplianceService } from './services/ComplianceRegisterService';
import ComplianceListView from './components/ComplianceListView';
import ComplianceForm from './components/ComplianceForm';
import * as JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IComplianceRegisterModuleProps {
  context: WebPartContext;
}

const ComplianceRegisterModule: React.FC<IComplianceRegisterModuleProps> = ({ context }) => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const service = useMemo(() => new ComplianceService(siteUrl), [siteUrl]);

  const [viewMode, setViewMode] = useState<'list' | 'form'>('list');
  const [items, setItems] = useState<IComplianceItem[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize] = useState(15);
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [listLoading, setListLoading] = useState(true);
  const [editingItem, setEditingItem] = useState<Partial<IComplianceItem> | null>(null);
  const [lookups, setLookups] = useState<ILookupSets>({
    complianceTypes: [], businesses: [], employees: [], projects: [], subcontractors: [], workers: []
  });
  const [lookupsLoaded, setLookupsLoaded] = useState(false);
  const [attachments, setAttachments] = useState<IAttachment[]>([]);
  const [pendingAttachments, setPendingAttachments] = useState<File[]>([]);
  const [attachmentsLoading, setAttachmentsLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [sortColumn, setSortColumn] = useState('Id');
  const [isAscending, setIsAscending] = useState(false);

  // ─── Data Loading ─────────────────────────────────────────────────────────

  const loadListData = async (): Promise<void> => {
    setListLoading(true);
    setErrorMessage('');
    try {
      const skip = (currentPage - 1) * pageSize;
      const result = await service.getItems(pageSize, skip, filters, sortColumn, isAscending);
      setItems(result.items);
      setTotalCount(result.totalCount);
      setListLoading(false);
    } catch (err) {
      setListLoading(false);
      setErrorMessage(`Failed to load data: ${(err as Error).message}`);
    }
  };

  const loadLookups = async (): Promise<void> => {
    try {
      const result = await service.getAllLookups();
      setLookups(result);
      setLookupsLoaded(true);
    } catch (err) {
      setErrorMessage(`Failed to load lookups: ${(err as Error).message}`);
    }
  };

  const loadAttachments = async (itemId: number): Promise<void> => {
    setAttachmentsLoading(true);
    try {
      const result = await service.getAttachments(itemId);
      setAttachments(result);
      setAttachmentsLoading(false);
    } catch {
      setAttachments([]);
      setAttachmentsLoading(false);
    }
  };

  useEffect(() => {
    if (!lookupsLoaded) {
      void loadLookups();
    }
  }, []);

  useEffect(() => {
    void loadListData();
  }, [currentPage, filters, sortColumn, isAscending]);

  // ─── Navigation ──────────────────────────────────────────────────────────

  const openNewForm = (): void => {
    setViewMode('form');
    setEditingItem(null);
    setAttachments([]);
    setPendingAttachments([]);
    setErrorMessage('');
  };

  const openEditForm = async (id: number): Promise<void> => {
    setListLoading(true);
    setErrorMessage('');
    try {
      const [item, atts] = await Promise.all([
        service.getItem(id),
        service.getAttachments(id)
      ]);
      setViewMode('form');
      setEditingItem(item);
      setAttachments(atts);
      setPendingAttachments([]);
      setListLoading(false);
      setAttachmentsLoading(false);
    } catch (err) {
      setListLoading(false);
      setErrorMessage(`Load error: ${(err as Error).message}`);
    }
  };

  const closeForm = (): void => {
    setViewMode('list');
    setEditingItem(null);
    setAttachments([]);
  };

  // ─── CRUD ─────────────────────────────────────────────────────────────────

  const flushPendingAttachments = async (itemId: number): Promise<void> => {
    if (pendingAttachments.length > 0) {
      for (const file of pendingAttachments) {
        await service.uploadAttachment(itemId, file);
      }
      setPendingAttachments([]);
    }
  };

  const handleSave = async (item: Partial<IComplianceItem>): Promise<void> => {
    setSaving(true);
    setErrorMessage('');
    try {
      if (editingItem?.Id) {
        await service.updateItem(editingItem.Id, item);
        await flushPendingAttachments(editingItem.Id);
        const updated = await service.getItem(editingItem.Id);
        setSaving(false);
        setEditingItem(updated);
        await loadAttachments(editingItem.Id);
      } else {
        const newId = await service.createItem(item);
        await flushPendingAttachments(newId);
        const created = await service.getItem(newId);
        setSaving(false);
        setEditingItem(created);
        await loadAttachments(newId);
      }
      await loadListData();
    } catch (err) {
      setSaving(false);
      setErrorMessage(`Save failed: ${(err as Error).message}`);
    }
  };

  const handleSaveNew = async (item: Partial<IComplianceItem>): Promise<void> => {
    setSaving(true);
    setErrorMessage('');
    try {
      const newId = await service.createItem(item);
      await flushPendingAttachments(newId);
      setSaving(false);
      setEditingItem(null);
      setAttachments([]);
      setPendingAttachments([]);
      await loadListData();
    } catch (err) {
      setSaving(false);
      setErrorMessage(`Save failed: ${(err as Error).message}`);
    }
  };

  const handleSaveClose = async (item: Partial<IComplianceItem>): Promise<void> => {
    setSaving(true);
    setErrorMessage('');
    try {
      if (editingItem?.Id) {
        await service.updateItem(editingItem.Id, item);
        await flushPendingAttachments(editingItem.Id);
      } else {
        const newId = await service.createItem(item);
        await flushPendingAttachments(newId);
      }
      setSaving(false);
      setViewMode('list');
      setEditingItem(null);
      setAttachments([]);
      setPendingAttachments([]);
      await loadListData();
    } catch (err) {
      setSaving(false);
      setErrorMessage(`Save failed: ${(err as Error).message}`);
    }
  };

  const handleDelete = async (id: number): Promise<void> => {
    if (!window.confirm('Are you sure you want to delete this record?')) return;
    try {
      await service.deleteItem(id);
      await loadListData();
    } catch (err) {
      setErrorMessage(`Delete failed: ${(err as Error).message}`);
    }
  };

  // ─── Attachments ─────────────────────────────────────────────────────────

  const handleUploadAttachment = async (file: File): Promise<void> => {
    if (!editingItem?.Id) {
      setPendingAttachments(prev => [...prev, file]);
      return;
    }
    try {
      await service.uploadAttachment(editingItem.Id, file);
      await loadAttachments(editingItem.Id);
    } catch (err) {
      setErrorMessage(`Upload failed: ${(err as Error).message}`);
    }
  };

  const handleRemoveAttachment = async (fileName: string): Promise<void> => {
    if (!editingItem?.Id) {
      setPendingAttachments(prev => prev.filter(f => f.name !== fileName));
      return;
    }
    if (!window.confirm(`Remove attachment "${fileName}"?`)) return;
    try {
      await service.deleteAttachment(editingItem.Id, fileName);
      await loadAttachments(editingItem.Id);
    } catch (err) {
      setErrorMessage(`Remove failed: ${(err as Error).message}`);
    }
  };

  const handleDownloadZip = async (item: IComplianceItem): Promise<void> => {
    try {
      setListLoading(true);
      setErrorMessage('');

      const [fullItem, createdBy] = await Promise.all([
        service.getItem(item.Id),
        service.getItemAuthor(item.Id)
      ]);

      const zip = new JSZip();
      const dateStr = new Date().toISOString().replace(/[:.]/g, '').slice(0, 14);
      const folderName = `ComplianceRecords${dateStr}`;
      const rootFolder = zip.folder(folderName);
      if (!rootFolder) throw new Error("Could not create ZIP folder");

      const headers = [
        'ID', 'Document #', 'Training For', 'Name', 'Position', 'Company',
        'Entity Active', 'Training Type', 'Booking Date', 'Booked With',
        'Issuing Authority', 'Issue Date', 'Expiry Date', 'Business Profile',
        'Date Entered', 'Entered By', 'Status', 'Expired', 'Days Remaining', 'Notes'
      ];

      const extractFirstPart = (text: string): string => text ? text.split(' | ')[0].trim() : '';
      const getName = (i: IComplianceItem): string => {
        switch (i.ComplianceFor) {
          case 'Business':      return i.BusinessName;
          case 'Employee':      return extractFirstPart(i.EmployeeName);
          case 'Project':       return extractFirstPart(i.ProjectTitle);
          case 'Subcontractor': return extractFirstPart(i.SubcontractorName);
          case 'Worker':        return extractFirstPart(i.WorkerName);
          default: return '';
        }
      };

      const getCompanyName = (i: IComplianceItem): string => {
        if (i.ComplianceFor === 'Worker') {
          const parts = (i.WorkerName || '').split(' | ');
          return parts[2]?.trim() || i.BusinessName || '';
        }
        if (i.ComplianceFor === 'Subcontractor') {
          return (i.SubcontractorName || '').split(' | ')[0].trim() || '';
        }
        if (i.MainBusinessProfileId && lookups.businesses) {
          const found = lookups.businesses.find(b => b.key === i.MainBusinessProfileId);
          if (found) return found.text;
        }
        return i.BusinessName || '';
      };

      const getTypeName = (id: number | null): string => {
        if (!id) return '';
        const found = lookups.complianceTypes.find(t => t.key === id);
        return found ? found.text : `Type ${id}`;
      };

      const formatDate = (dateStr: string | null | undefined): string => {
        if (!dateStr) return '';
        try {
          const d = new Date(dateStr);
          return d.toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' });
        } catch { return dateStr; }
      };

      const row = [
        fullItem.Id,
        fullItem.DocumentNumber || '',
        fullItem.ComplianceFor,
        getName(fullItem),
        fullItem.Position || '',
        getCompanyName(fullItem),
        fullItem.EntityActive || 'Yes',
        getTypeName(fullItem.ComplianceTypeId),
        formatDate(fullItem.BookingDate),
        fullItem.BookedWith || '',
        fullItem.IssuingAuthority || '',
        formatDate(fullItem.IssueDate),
        formatDate(fullItem.ExpiryDate),
        fullItem.BusinessProfile || fullItem.BusinessName || '',
        formatDate(fullItem.DateEntered),
        createdBy || '',
        fullItem.Status || '',
        fullItem.Expired || '',
        fullItem.DaysRemaining || '',
        fullItem.Notes || ''
      ];

      const csvContent = [headers, row].map(r => r.map(c => `"${String(c ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');
      rootFolder.file(`${folderName}.csv`, csvContent);

      const atts = await service.getAttachments(fullItem.Id);
      if (atts.length > 0) {
        const attFolder = rootFolder.folder('Attachments');
        if (attFolder) {
          for (const att of atts) {
            const blob = await service.getAttachmentBlob(att.ServerRelativeUrl);
            attFolder.file(att.FileName, blob);
          }
        }
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      saveAs(zipBlob, `${folderName}.zip`);
      setListLoading(false);
    } catch (err) {
      setListLoading(false);
      setErrorMessage(`Download failed: ${(err as Error).message}`);
    }
  };

  const handleExport = (format: 'csv' | 'excel'): void => {
    const extractFirstPart = (text: string): string => text ? text.split(' | ')[0].trim() : '';

    const getCleanName = (item: IComplianceItem): string => {
      switch (item.ComplianceFor) {
        case 'Business':      return item.BusinessName;
        case 'Employee':      return extractFirstPart(item.EmployeeName);
        case 'Project':       return extractFirstPart(item.ProjectTitle);
        case 'Subcontractor': return extractFirstPart(item.SubcontractorName);
        case 'Worker':        return extractFirstPart(item.WorkerName);
        default: return '';
      }
    };

    const getCleanCompany = (item: IComplianceItem): string => {
      if (item.ComplianceFor === 'Worker') {
        const parts = (item.WorkerName || '').split(' | ');
        return parts[2]?.trim() || item.BusinessName || '';
      }
      if (item.ComplianceFor === 'Subcontractor') {
        return (item.SubcontractorName || '').split(' | ')[0].trim() || '';
      }
      if (item.MainBusinessProfileId && lookups.businesses) {
        const found = lookups.businesses.find(b => b.key === item.MainBusinessProfileId);
        if (found) return found.text;
      }
      return item.BusinessName || '';
    };

    const formatDate = (dateStr: string | null): string => {
      if (!dateStr) return '';
      const d = new Date(dateStr);
      return d.toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' });
    };

    const getTypeName = (id: number | null): string => {
      if (!id) return '';
      const found = lookups.complianceTypes.find(t => t.key === id);
      return found ? found.text : `Type ${id}`;
    };

    const data = items.map(i => ({
      'ID': i.Id,
      'Compliance For': i.ComplianceFor,
      'Name': getCleanName(i),
      'Company': getCleanCompany(i),
      'Entity': i.EntityActive || 'Yes',
      'Type': getTypeName(i.ComplianceTypeId),
      'Booking': i.IsBooking ? 'Yes' : 'No',
      'Booked With': i.BookedWith,
      'Expiry Date': formatDate(i.ExpiryDate),
      'Status': i.Status,
      'Expired': i.Expired,
      'Days Remaining': i.DaysRemaining
    }));

    if (format === 'csv') {
      const headers = Object.keys(data[0] || {});
      const rows = data.map(d => headers.map(h => {
        const val = (d as any)[h];
        return `"${String(val).replace(/"/g, '""')}"`;
      }).join(','));
      const csv = [headers.join(','), ...rows].join('\n');
      const blob = new Blob([csv], { type: 'text/csv' });
      saveAs(blob, 'ComplianceRegister.csv');
    } else {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Compliance");
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'ComplianceRegister.xlsx');
    }
  };

  const handlePageChange = (page: number): void => {
    setCurrentPage(page);
  };

  const handleFilterChange = (key: string, value: string): void => {
    setFilters(prev => ({ ...prev, [key]: value }));
    setCurrentPage(1);
  };

  const handleSort = (column: string): void => {
    setIsAscending(prev => (sortColumn === column ? !prev : true));
    setSortColumn(column);
    setCurrentPage(1);
  };

  const mergedAttachments = [
    ...attachments,
    ...(pendingAttachments || []).map(f => ({ FileName: f.name, ServerRelativeUrl: '' }))
  ];

  return (
    <div style={{ position: 'relative', width: '100%', minHeight: '400px' }}>
      {errorMessage && (
        <div style={{
          background: '#fee2e2', color: '#b91c1c', padding: '12px 16px', borderRadius: '8px',
          marginBottom: '16px', border: '1px solid #fecaca', display: 'flex', justifyContent: 'space-between'
        }}>
          <span>⚠ {errorMessage}</span>
          <button style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#b91c1c', fontWeight: 'bold' }}
            onClick={() => setErrorMessage('')}>✖</button>
        </div>
      )}

      {viewMode === 'list' && (
        <ComplianceListView
          items={items}
          totalCount={totalCount}
          currentPage={currentPage}
          pageSize={pageSize}
          loading={listLoading}
          filters={filters}
          siteUrl={siteUrl}
          lookups={lookups}
          sortColumn={sortColumn}
          isAscending={isAscending}
          onPageChange={handlePageChange}
          onFilterChange={handleFilterChange}
          onSort={handleSort}
          onNewClick={openNewForm}
          onEditClick={id => { void openEditForm(id); }}
          onDeleteClick={id => { void handleDelete(id); }}
          onRefresh={() => { void loadListData(); }}
          onExport={handleExport}
          onDownload={item => { void handleDownloadZip(item); }}
        />
      )}

      {viewMode === 'form' && (
        <ComplianceForm
          item={editingItem}
          lookups={lookups}
          attachments={mergedAttachments}
          attachmentsLoading={attachmentsLoading}
          saving={saving}
          siteUrl={siteUrl}
          onSave={handleSave}
          onSaveNew={handleSaveNew}
          onSaveClose={handleSaveClose}
          onClose={closeForm}
          onUploadAttachment={handleUploadAttachment}
          onRemoveAttachment={handleRemoveAttachment}
        />
      )}
    </div>
  );
};

export default ComplianceRegisterModule;
