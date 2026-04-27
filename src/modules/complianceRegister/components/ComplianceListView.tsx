import * as React from 'react';
import styles from './ComplianceListView.module.scss';
import { IComplianceItem, ILookupSets } from '../models/IComplianceItem';

interface IComplianceListViewProps {
  items: IComplianceItem[];
  totalCount: number;
  currentPage: number;
  pageSize: number;
  loading: boolean;
  filters: Record<string, string>;
  siteUrl: string;
  lookups: ILookupSets;
  onPageChange: (page: number) => void;
  onFilterChange: (key: string, value: string) => void;
  onNewClick: () => void;
  onEditClick: (id: number) => void;
  onDeleteClick: (id: number) => void;
  onRefresh: () => void;
  onExport: (format: 'csv' | 'excel') => void;
  onDownload: (item: IComplianceItem) => void;
  sortColumn: string;
  isAscending: boolean;
  onSort: (column: string) => void;
}

const ComplianceListView: React.FC<IComplianceListViewProps> = (props) => {
  const {
    items, totalCount, currentPage, pageSize, loading, filters, lookups,
    onPageChange, onFilterChange, onNewClick, onEditClick, onDeleteClick,
    onRefresh, onExport, onDownload, sortColumn, isAscending, onSort
  } = props;

  /* ─── state ────────────────────────────────────────────────────────────── */
  const [selectedIds, setSelectedIds]   = React.useState<Set<number>>(new Set());
  const [isExportMenuOpen, setExportMenuOpen] = React.useState(false);

  // Client-side filters for computed / lookup columns
  const [cfType,        setCfType]        = React.useState('');
  const [cfBookDate,    setCfBookDate]    = React.useState('');
  const [cfBookedWith,  setCfBookedWith]  = React.useState('');
  const [cfExpiryDate,  setCfExpiryDate]  = React.useState('');
  const [cfStatus,      setCfStatus]      = React.useState('');
  const [cfExpired,     setCfExpired]     = React.useState('');
  const [cfDays,        setCfDays]        = React.useState('');

  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

  /* ─── selection ─────────────────────────────────────────────────────────── */
  const toggleSelectAll = (): void => {
    setSelectedIds(selectedIds.size === items.length && items.length > 0
      ? new Set()
      : new Set(items.map(i => i.Id))
    );
  };
  const toggleSelect = (id: number): void => {
    const next = new Set(selectedIds);
    if (next.has(id)) next.delete(id); else next.add(id);
    setSelectedIds(next);
  };

  /* ─── display helpers ───────────────────────────────────────────────────── */
  const getDisplayName = (item: IComplianceItem): string => {
    const first = (s: string) => (s || '').split(' | ')[0].trim();
    switch (item.ComplianceFor) {
      case 'Business':      return item.BusinessName;
      case 'Employee':      return first(item.EmployeeName);
      case 'Project':       return first(item.ProjectTitle);
      case 'Subcontractor': return first(item.SubcontractorName);
      case 'Worker':        return first(item.WorkerName);
      default: return '';
    }
  };

  const getCompanyName = (item: IComplianceItem): string => {
    if (item.ComplianceFor === 'Worker') {
      const p = (item.WorkerName || '').split(' | ');
      return p[2]?.trim() || item.BusinessName || '';
    }
    if (item.ComplianceFor === 'Subcontractor') return (item.SubcontractorName || '').split(' | ')[0].trim();
    if (item.MainBusinessProfileId) {
      const b = lookups.businesses.find(x => x.key === item.MainBusinessProfileId);
      if (b) return b.text;
    }
    return item.BusinessName || '';
  };

  const getTypeName = (id: number | null): string => {
    if (!id) return '';
    const t = lookups.complianceTypes.find(x => x.key === id);
    return t ? t.text : `Type ${id}`;
  };

  const formatDate = (s: string | null): string => {
    if (!s) return '';
    return new Date(s).toLocaleDateString('en-AU', { day: '2-digit', month: '2-digit', year: 'numeric' });
  };

  const getCfClass = (type: string): string => {
    const m: Record<string, string> = {
      Employee: styles.cfEmployee, Business: styles.cfBusiness,
      Project: styles.cfProject, Subcontractor: styles.cfSubcontractor, Worker: styles.cfWorker,
    };
    return m[type] || styles.cfEmployee;
  };

  const getCfIcon = (type: string): string => {
    const ic: Record<string, string> = { Business: '🏢', Employee: '👤', Project: '📋', Subcontractor: '🔧', Worker: '👷' };
    return ic[type] || '📄';
  };

  const getDaysClass = (val: string | undefined): string => {
    const n = parseInt(val || '', 10);
    if (isNaN(n)) return '';
    if (n < 0) return styles.daysDanger;
    if (n <= 30) return styles.daysWarning;
    return styles.daysGood;
  };

  /* ─── client-side filtering ─────────────────────────────────────────────── */
  const filteredItems = items.filter(item => {
    if (cfType       && !getTypeName(item.ComplianceTypeId).toLowerCase().includes(cfType.toLowerCase()))     return false;
    if (cfBookDate   && !formatDate(item.BookingDate).toLowerCase().includes(cfBookDate.toLowerCase()))       return false;
    if (cfBookedWith  && !(item.BookedWith  || '').toLowerCase().includes(cfBookedWith.toLowerCase()))         return false;
    if (cfExpiryDate  && !formatDate(item.ExpiryDate).toLowerCase().includes(cfExpiryDate.toLowerCase()))     return false;
    if (cfStatus      && !(item.Status    || '').toLowerCase().includes(cfStatus.toLowerCase()))               return false;
    if (cfExpired    && !(item.Expired   || '').toLowerCase().includes(cfExpired.toLowerCase()))              return false;
    if (cfDays       && !String(item.DaysRemaining || '').includes(cfDays))                                  return false;
    return true;
  });

  /* ─── pagination ────────────────────────────────────────────────────────── */
  const getPaginationPages = (): (number | string)[] => {
    const pages: (number | string)[] = [];
    if (totalPages <= 8) {
      for (let i = 1; i <= totalPages; i++) pages.push(i);
    } else {
      pages.push(1, 2, 3, 4, 5, 6);
      if (currentPage > 7) pages.push('...');
      if (currentPage > 6 && currentPage < totalPages - 1) pages.push(currentPage);
      pages.push('...', totalPages - 1, totalPages);
    }
    return pages;
  };

  const firstItem  = totalCount === 0 ? 0 : (currentPage - 1) * pageSize + 1;
  const lastItem   = Math.min(currentPage * pageSize, totalCount);
  const sortArrow  = (col: string) => sortColumn === col ? (isAscending ? ' ▲' : ' ▼') : ' ↕';

  /* ─── render ────────────────────────────────────────────────────────────── */
  return (
    <div className={styles.listViewContainer}>

      {/* ── Toolbar ── */}
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          <div className={styles.toolbarIconBox}>📋</div>
          <div className={styles.toolbarTitleBlock}>
            <span className={styles.toolbarTitle}>Compliance Register</span>
            <span className={styles.toolbarSubtitle}>Manage and track all compliance records</span>
          </div>
        </div>
        <div className={styles.toolbarActions}>
          <button className={`${styles.toolbarBtn} ${styles.btnNew}`} onClick={onNewClick} id="btn-new-compliance">
            <span className={styles.btnIcon}>✚</span> New
          </button>
          <button
            className={`${styles.toolbarBtn} ${styles.btnDownload}`}
            onClick={() => { const f = items.find(i => selectedIds.has(i.Id)); if (f) onDownload(f); }}
            id="btn-download-compliance"
          >
            <span className={styles.btnIcon}>⬇</span> Download
          </button>
          <button
            className={`${styles.toolbarBtn} ${styles.btnDelete}`}
            onClick={() => { selectedIds.forEach(id => onDeleteClick(id)); setSelectedIds(new Set()); }}
            id="btn-delete-compliance"
          >
            <span className={styles.btnIcon}>🗑</span> Delete
          </button>
          <div className={styles.exportMenuContainer}>
            <button
              className={`${styles.toolbarBtn} ${styles.btnExport}`}
              onClick={() => setExportMenuOpen(!isExportMenuOpen)}
              id="btn-export-compliance"
            >
              <span className={styles.btnIcon}>📤</span> Export<span className={styles.dropdownArrow}> ▾</span>
            </button>
            {isExportMenuOpen && (
              <div className={styles.exportMenu}>
                <div className={styles.exportMenuItem} onClick={() => { onExport('csv');   setExportMenuOpen(false); }}>📄 Export as CSV</div>
                <div className={styles.exportMenuItem} onClick={() => { onExport('excel'); setExportMenuOpen(false); }}>📊 Export as Excel</div>
              </div>
            )}
          </div>
          <button className={`${styles.toolbarBtn} ${styles.btnRefresh}`} onClick={onRefresh} id="btn-refresh-compliance">
            <span className={styles.btnIcon}>🔄</span> Refresh
          </button>
        </div>
      </div>

      {/* ── Summary Strip ── */}
      <div className={styles.summaryBar}>
        <span className={styles.summaryChip}>📊 {totalCount} Records</span>
        <span className={styles.summaryHint}>
          {selectedIds.size > 0 ? `${selectedIds.size} row${selectedIds.size > 1 ? 's' : ''} selected` : 'Click an ID to edit — check rows to select'}
        </span>
      </div>

      {/* ── Grid ── */}
      <div className={styles.gridWrapper}>
        {loading && <div className={styles.loadingOverlay}><div className={styles.spinner} /></div>}
        <table className={styles.grid}>
          <thead>
            {/* Column header row */}
            <tr className={styles.headerRow}>
              <th className={`${styles.th} ${styles.checkCol}`}>
                <input type="checkbox" checked={selectedIds.size === items.length && items.length > 0} onChange={toggleSelectAll} id="chk-select-all" />
              </th>
              <th className={`${styles.th} ${styles.iconCol}`}></th>
              <th className={`${styles.th} ${styles.sortableHeader}`} onClick={() => onSort('Id')}>ID{sortArrow('Id')}</th>
              <th className={styles.th}>Compliance For <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Name <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Company <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Entity Active</th>
              <th className={styles.th}>Type <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Booking Date <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Booked With <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Expiry Date <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Status <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Expired <span className={styles.filterIcon}>▼</span></th>
              <th className={styles.th}>Days Remaining <span className={styles.filterIcon}>▼</span></th>
            </tr>

            {/* Filter row — every column now has a search input */}
            <tr className={styles.filterRow}>
              <td className={styles.filterTd}></td>
              <td className={styles.filterTd}></td>

              {/* ID — server-side OData */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search…" value={filters['Id'] || ''}
                  onChange={e => onFilterChange('Id', e.target.value)} id="filter-id" />
              </td>

              {/* Compliance For — server-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search…" value={filters['ComplianceFor'] || ''}
                  onChange={e => onFilterChange('ComplianceFor', e.target.value)} id="filter-compliance-for" />
              </td>

              {/* Name — server-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search…" value={filters['Name'] || ''}
                  onChange={e => onFilterChange('Name', e.target.value)} id="filter-name" />
              </td>

              {/* Company — server-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search…" value={filters['Company'] || ''}
                  onChange={e => onFilterChange('Company', e.target.value)} id="filter-company" />
              </td>

              {/* Entity Active — no filter */}
              <td className={styles.filterTd}></td>

              {/* Type — client-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search type…" value={cfType}
                  onChange={e => setCfType(e.target.value)} id="filter-type" />
              </td>

              {/* Booking Date — client-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="dd/mm/yyyy" value={cfBookDate}
                  onChange={e => setCfBookDate(e.target.value)} id="filter-booking-date" />
              </td>

              {/* Booked With — client-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Search…" value={cfBookedWith}
                  onChange={e => setCfBookedWith(e.target.value)} id="filter-booked-with" />
              </td>

              {/* Expiry Date — client-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="dd/mm/yyyy" value={cfExpiryDate}
                  onChange={e => setCfExpiryDate(e.target.value)} id="filter-expiry-date" />
              </td>

              {/* Status — client-side (type 'active' or 'expired') */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Active / Expired" value={cfStatus}
                  onChange={e => setCfStatus(e.target.value)} id="filter-status" />
              </td>

              {/* Expired — client-side (type 'yes' or 'no') */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="Yes / No" value={cfExpired}
                  onChange={e => setCfExpired(e.target.value)} id="filter-expired" />
              </td>

              {/* Days Remaining — client-side */}
              <td className={styles.filterTd}>
                <input className={styles.filterInput} placeholder="e.g. 30" value={cfDays}
                  onChange={e => setCfDays(e.target.value)} id="filter-days-remaining" />
              </td>
            </tr>
          </thead>

          <tbody>
            {/* Empty state */}
            {filteredItems.length === 0 && !loading && (
              <tr>
                <td colSpan={14} className={styles.emptyRow}>
                  <div className={styles.emptyBox}>
                    <span className={styles.emptyIcon}>📂</span>
                    <span className={styles.emptyTitle}>No compliance records found</span>
                    <span className={styles.emptySub}>Try adjusting your search filters or click <strong>New</strong> to add a record</span>
                  </div>
                </td>
              </tr>
            )}

            {/* Data rows (filtered client-side) */}
            {filteredItems.map(item => {
              const isSelected = selectedIds.has(item.Id);
              return (
                <tr key={item.Id} className={`${styles.dataRow} ${isSelected ? styles.selectedRow : ''}`}>
                  <td className={`${styles.td} ${styles.checkCol}`}>
                    <input type="checkbox" checked={isSelected} onChange={() => toggleSelect(item.Id)} id={`chk-row-${item.Id}`} />
                  </td>
                  <td className={`${styles.td} ${styles.iconCol}`}>
                    {item.HasAttachments && (
                      <span className={styles.attachIcon} title="Has attachments — click to download" onClick={() => onDownload(item)}>📎</span>
                    )}
                  </td>
                  <td className={styles.td}>
                    <span className={styles.idPill} onClick={() => onEditClick(item.Id)} id={`link-id-${item.Id}`}>{item.Id}</span>
                  </td>
                  <td className={styles.td}>
                    <span className={getCfClass(item.ComplianceFor)}>{getCfIcon(item.ComplianceFor)} {item.ComplianceFor}</span>
                  </td>
                  <td className={styles.td}>{getDisplayName(item)}</td>
                  <td className={styles.td}>{getCompanyName(item)}</td>
                  <td className={styles.td}><span className={styles.yesTag}>✔ Yes</span></td>
                  <td className={styles.td}>{getTypeName(item.ComplianceTypeId)}</td>
                  <td className={styles.td}>{formatDate(item.BookingDate)}</td>
                  <td className={styles.td}>{item.BookedWith}</td>
                  <td className={styles.td}>{formatDate(item.ExpiryDate)}</td>
                  <td className={styles.td}>
                    {item.Status === 'Active'
                      ? <span className={styles.statusActive}>Active</span>
                      : item.Status === 'Expired'
                        ? <span className={styles.statusExpired}>Expired</span>
                        : item.Status}
                  </td>
                  <td className={styles.td}>{item.Expired}</td>
                  <td className={styles.td}>
                    <span className={getDaysClass(item.DaysRemaining)}>
                      {item.DaysRemaining !== undefined && item.DaysRemaining !== '' ? `${item.DaysRemaining}d` : ''}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* ── Pagination ── */}
      <div className={styles.pagination}>
        <span className={styles.pageInfo}>
          Page {currentPage} of {totalPages} ({totalCount} items) — Showing {firstItem}–{lastItem}
        </span>
        <div className={styles.pageNav}>
          <button className={styles.pageBtn} disabled={currentPage === 1}         onClick={() => onPageChange(1)}              id="page-first">«</button>
          <button className={styles.pageBtn} disabled={currentPage === 1}         onClick={() => onPageChange(currentPage - 1)} id="page-prev">‹</button>
          {getPaginationPages().map((p, idx) =>
            typeof p === 'number' ? (
              <button key={idx} className={`${styles.pageBtn} ${p === currentPage ? styles.activePageBtn : ''}`}
                onClick={() => onPageChange(p)} id={`page-${p}`}>{p}</button>
            ) : (
              <span key={idx} className={styles.pageDots}>…</span>
            )
          )}
          <button className={styles.pageBtn} disabled={currentPage === totalPages} onClick={() => onPageChange(currentPage + 1)} id="page-next">›</button>
          <button className={styles.pageBtn} disabled={currentPage === totalPages} onClick={() => onPageChange(totalPages)}      id="page-last">»</button>
        </div>
      </div>

    </div>
  );
};

export default ComplianceListView;
