import * as React from 'react';
import styles from './ComplianceForm.module.scss';
import { IComplianceItem, ILookupOption, ILookupSets, ComplianceForType } from '../models/IComplianceItem';
import { IAttachment } from '../models/IComplianceItem';

interface IComplianceFormProps {
  item: Partial<IComplianceItem> | null;  // null = New, populated = Edit
  lookups: ILookupSets;
  attachments: IAttachment[];
  attachmentsLoading: boolean;
  saving: boolean;
  siteUrl: string;
  onSave: (item: Partial<IComplianceItem>) => Promise<void>;
  onSaveNew: (item: Partial<IComplianceItem>) => Promise<void>;
  onSaveClose: (item: Partial<IComplianceItem>) => Promise<void>;
  onClose: () => void;
  onUploadAttachment: (file: File) => Promise<void>;
  onRemoveAttachment: (fileName: string) => Promise<void>;
}

interface IFormState {
  complianceFor: ComplianceForType;
  businessId: number | null;
  employeeId: number | null;
  projectId: number | null;
  subcontractorId: number | null;
  workerId: number | null;
  complianceTypeId: number | null;
  isBooking: boolean;
  bookingDate: string;
  bookedWith: string;
  documentNumber: string;
  issuingAuthority: string;
  issueDate: string;
  renewalNotRequired: boolean;
  expiryDate: string;
  notes: string;
  mainBusinessProfileId: number | null;
}

interface IValidationErrors {
  complianceFor?: string;
  complianceTypeId?: string;
  nameField?: string;
  mainBusinessProfileId?: string;
}

const COMPLIANCE_FOR_OPTIONS: { value: ComplianceForType; icon: string }[] = [
  { value: 'Business',      icon: '🏢' },
  { value: 'Employee',      icon: '👤' },
  { value: 'Project',       icon: '📋' },
  { value: 'Subcontractor', icon: '🔧' },
  { value: 'Worker',        icon: '👷' },
];

const emptyState = (): IFormState => ({
  complianceFor: '',
  businessId: null,
  employeeId: null,
  projectId: null,
  subcontractorId: null,
  workerId: null,
  complianceTypeId: null,
  isBooking: false,
  bookingDate: '',
  bookedWith: '',
  documentNumber: '',
  issuingAuthority: '',
  issueDate: '',
  renewalNotRequired: false,
  expiryDate: '',
  notes: '',
  mainBusinessProfileId: null
});

const itemToState = (item: Partial<IComplianceItem>): IFormState => ({
  complianceFor: item.ComplianceFor || '',
  businessId: item.BusinessId || null,
  employeeId: item.EmployeeId || null,
  projectId: item.ProjectId || null,
  subcontractorId: item.SubcontractorId || null,
  workerId: item.WorkerId || null,
  complianceTypeId: item.ComplianceTypeId || null,
  isBooking: item.IsBooking || false,
  bookingDate: item.BookingDate ? item.BookingDate.split('T')[0] : '',
  bookedWith: item.BookedWith || '',
  documentNumber: item.DocumentNumber || '',
  issuingAuthority: item.IssuingAuthority || '',
  issueDate: item.IssueDate ? item.IssueDate.split('T')[0] : '',
  renewalNotRequired: item.RenewalNotRequired || false,
  expiryDate: item.ExpiryDate ? item.ExpiryDate.split('T')[0] : '',
  notes: item.Notes || '',
  mainBusinessProfileId: item.MainBusinessProfileId || null
});

const stateToPayload = (state: IFormState): Partial<IComplianceItem> => ({
  ComplianceFor: state.complianceFor,
  ComplianceTypeId: state.complianceTypeId,
  IsBooking: state.isBooking,
  BookingDate: state.bookingDate ? `${state.bookingDate}T00:00:00Z` : null,
  BookedWith: state.bookedWith,
  DocumentNumber: state.documentNumber,
  IssuingAuthority: state.issuingAuthority,
  IssueDate: state.issueDate ? `${state.issueDate}T00:00:00Z` : null,
  RenewalNotRequired: state.renewalNotRequired,
  ExpiryDate: state.expiryDate ? `${state.expiryDate}T00:00:00Z` : null,
  Notes: state.notes,
  BusinessId: state.complianceFor === 'Business' ? state.businessId : null,
  EmployeeId: state.complianceFor === 'Employee' ? state.employeeId : null,
  ProjectId: state.complianceFor === 'Project' ? state.projectId : null,
  SubcontractorId: state.complianceFor === 'Subcontractor' ? state.subcontractorId : null,
  WorkerId: state.complianceFor === 'Worker' ? state.workerId : null,
  MainBusinessProfileId: state.mainBusinessProfileId
});

// ── Inline Attachment Panel (modernised) ─────────────────────────────────────
interface IAttachmentPanelInlineProps {
  attachments: IAttachment[];
  siteUrl: string;
  loading: boolean;
  onUpload: (file: File) => Promise<void>;
  onRemove: (fileName: string) => Promise<void>;
}

const AttachmentPanelInline: React.FC<IAttachmentPanelInlineProps> = ({
  attachments, siteUrl, loading, onUpload, onRemove
}) => {
  const [selected, setSelected] = React.useState<string | null>(null);
  const [uploading, setUploading] = React.useState(false);
  const fileRef = React.useRef<HTMLInputElement>(null);

  const handleUpload = (): void => { if (fileRef.current) fileRef.current.click(); };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try { await onUpload(file); } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = '';
    }
  };

  const handleRemove = async (): Promise<void> => {
    if (!selected) return;
    await onRemove(selected);
    setSelected(null);
  };

  const handleDownload = (): void => {
    if (!selected) return;
    const att = attachments.find(a => a.FileName === selected);
    if (!att) return;
    if (att.ServerRelativeUrl.startsWith('blob:')) {
      window.open(att.ServerRelativeUrl, '_blank');
    } else {
      window.open(`${siteUrl}${att.ServerRelativeUrl}`, '_blank');
    }
  };

  return (
    <div className={styles.attachCard}>
      <div className={styles.attachHeader}>
        <div className={styles.attachHeaderAccent} />
        <span className={styles.attachHeaderTitle}>📎 Attachments</span>
        <span className={styles.attachCount}>{attachments.length}</span>
      </div>
      <div className={styles.attachBody}>
        <div className={styles.attachActions}>
          <button className={styles.attachBtnUpload} onClick={handleUpload} disabled={uploading} id="btn-upload-attachment">
            {uploading ? '⏳ Uploading…' : '⬆ Upload'}
          </button>
          <button className={styles.attachBtnSecondary} onClick={handleRemove} disabled={!selected} id="btn-remove-attachment">
            🗑 Remove
          </button>
          <button className={styles.attachBtnSecondary} onClick={handleDownload} disabled={!selected} id="btn-download-attachment">
            ⬇ Download
          </button>
          <input ref={fileRef} type="file" style={{ display: 'none' }} onChange={handleFileChange} id="file-input-hidden" />
        </div>

        {loading && <div className={styles.loadingText}>Loading attachments…</div>}

        {!loading && attachments.length === 0 && (
          <div className={styles.emptyAttach}>
            <span className={styles.emptyAttachIcon}>📂</span>
            <span>No attachments yet</span>
            <span style={{ fontSize: '11.5px', color: '#b0bec5' }}>Upload files using the button above</span>
          </div>
        )}

        {!loading && attachments.length > 0 && (
          <ul className={styles.fileList}>
            {attachments.map(att => (
              <li
                key={att.FileName}
                className={`${styles.fileItem} ${selected === att.FileName ? styles.fileItemSelected : ''}`}
                onClick={() => setSelected(att.FileName === selected ? null : att.FileName)}
                id={`attach-${att.FileName.replace(/[^a-zA-Z0-9]/g, '-')}`}
              >
                <span className={styles.fileIcon}>📄</span>
                <span className={styles.fileName}>{att.FileName}</span>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
};

// ── Main Form ────────────────────────────────────────────────────────────────
const ComplianceForm: React.FC<IComplianceFormProps> = ({
  item, lookups, attachments, attachmentsLoading, saving,
  siteUrl, onSave, onSaveNew, onSaveClose, onClose,
  onUploadAttachment, onRemoveAttachment
}) => {
  const isNew = !item || !item.Id;
  const [form, setForm] = React.useState<IFormState>(() => item ? itemToState(item) : emptyState());
  const [errors, setErrors] = React.useState<IValidationErrors>({});

  React.useEffect(() => {
    setForm(item ? itemToState(item) : emptyState());
    setErrors({});
  }, [item]);

  const set = <K extends keyof IFormState>(key: K, value: IFormState[K]): void => {
    setForm(prev => ({ ...prev, [key]: value }));
    if (errors[key as keyof IValidationErrors]) {
      setErrors(prev => ({ ...prev, [key]: undefined }));
    }
  };

  const handleIsBookingChange = (checked: boolean): void => {
    setForm(prev => ({
      ...prev,
      isBooking: checked,
      ...(checked
        ? { documentNumber: '', issuingAuthority: '', issueDate: '', renewalNotRequired: false, expiryDate: '' }
        : { bookingDate: '', bookedWith: '' }
      )
    }));
  };

  const handleComplianceForChange = (value: ComplianceForType): void => {
    setForm(prev => ({
      ...prev,
      complianceFor: value,
      businessId: null, employeeId: null, projectId: null,
      subcontractorId: null, workerId: null
    }));
  };

  const getDynamicValue = (): number | null => {
    switch (form.complianceFor) {
      case 'Business': return form.businessId;
      case 'Employee': return form.employeeId;
      case 'Project': return form.projectId;
      case 'Subcontractor': return form.subcontractorId;
      case 'Worker': return form.workerId;
      default: return null;
    }
  };

  const validate = (): boolean => {
    const newErrors: IValidationErrors = {};
    if (!form.mainBusinessProfileId) newErrors.mainBusinessProfileId = 'Business Profile is required';
    if (!form.complianceFor) newErrors.complianceFor = 'Compliance For is required';
    if (!form.complianceTypeId) newErrors.complianceTypeId = 'Compliance Type is required';
    if (form.complianceFor && !getDynamicValue()) newErrors.nameField = `${form.complianceFor} is required`;
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const getDynamicOptions = (): ILookupOption[] => {
    let arr: ILookupOption[] = [];
    switch (form.complianceFor) {
      case 'Business': arr = lookups.businesses; break;
      case 'Employee': arr = lookups.employees; break;
      case 'Project': arr = lookups.projects; break;
      case 'Subcontractor': arr = lookups.subcontractors; break;
      case 'Worker': arr = lookups.workers; break;
      default: return [];
    }
    let replaceWith = 'Supply Workforce';
    if (form.mainBusinessProfileId) {
      const match = lookups.businesses.find(b => b.key === form.mainBusinessProfileId);
      if (match) replaceWith = match.text;
    }
    return arr.map(opt => ({
      ...opt,
      text: typeof opt.text === 'string' && form.complianceFor !== 'Employee' && opt.text.includes('Supply Workforce')
        ? opt.text.replace('Supply Workforce', replaceWith)
        : opt.text
    }));
  };

  const setDynamicValue = (id: number | null): void => {
    switch (form.complianceFor) {
      case 'Business': set('businessId', id); break;
      case 'Employee': set('employeeId', id); break;
      case 'Project': set('projectId', id); break;
      case 'Subcontractor': set('subcontractorId', id); break;
      case 'Worker': set('workerId', id); break;
    }
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    await onSave(stateToPayload(form));
  };

  const handleSaveNew = async (): Promise<void> => {
    if (!validate()) return;
    await onSaveNew(stateToPayload(form));
    setForm(emptyState());
    setErrors({});
  };

  const handleSaveClose = async (): Promise<void> => {
    if (!validate()) return;
    await onSaveClose(stateToPayload(form));
  };

  const showDynamicField = form.complianceFor !== '';
  const showBookingFields = form.isBooking;
  const showDocFields = !form.isBooking;
  const showExpiryDate = !form.isBooking && !form.renewalNotRequired;

  return (
    <div className={styles.formContainer}>

      {/* ── Header ── */}
      <div className={styles.formHeader}>
        <div className={styles.formTitleGroup}>
          <div className={styles.formTitleIcon}>📝</div>
          <div className={styles.formTitleText}>
            <span className={styles.formTitle}>
              {isNew ? 'New Compliance Record' : `Compliance #${item?.Id}`}
            </span>
            <span className={styles.formSubtitle}>
              {isNew ? 'Fill in the details below to create a new record' : 'Review and update the compliance details'}
            </span>
          </div>
        </div>

        <div className={styles.headerActions}>
          <button className={styles.btnSave} onClick={handleSave} disabled={saving} id="btn-form-save">
            💾 Save
          </button>
          <button className={styles.btnSaveNew} onClick={handleSaveNew} disabled={saving} id="btn-form-save-new">
            ➕ Save &amp; New
          </button>
          <button className={styles.btnSaveClose} onClick={handleSaveClose} disabled={saving} id="btn-form-save-close">
            ✔ Save &amp; Close
          </button>
          <button className={styles.btnClose} onClick={onClose} id="btn-form-close">
            ✖ Close
          </button>
        </div>
      </div>

      {/* Saving bar */}
      {saving && (
        <div className={styles.savingBar}>
          <div className={styles.savingSpinner} />
          <span>Saving your changes, please wait…</span>
        </div>
      )}

      {/* ── Body ── */}
      <div className={styles.formBody}>

        {/* LEFT — Form Fields */}
        <div className={styles.leftPanel}>

          {/* ▸ Card 1: Entity Configuration */}
          <div className={styles.card}>
            <div className={styles.cardHeader}>
              <div className={styles.cardHeaderAccent} />
              <span className={styles.cardHeaderTitle}>Entity Configuration</span>
              <span className={styles.cardHeaderIcon}>🏛</span>
            </div>
            <div className={styles.cardBody}>

              {/* Business Profile */}
              <div className={styles.formRow}>
                <label className={styles.formLabel}>
                  Business Profile <span className={styles.requiredStar}>*</span>
                </label>
                <div className={errors.mainBusinessProfileId ? styles.inputError : ''}>
                  <div className={styles.selectWrapper}>
                    <select
                      className={styles.select}
                      value={form.mainBusinessProfileId || ''}
                      onChange={e => setForm({ ...form, mainBusinessProfileId: e.target.value ? Number(e.target.value) : null })}
                    >
                      <option value="">— Select Business Profile —</option>
                      {lookups.businesses.map(opt => (
                        <option key={opt.key} value={opt.key}>{opt.text}</option>
                      ))}
                    </select>
                  </div>
                  {errors.mainBusinessProfileId && <span className={styles.errorMsg}>{errors.mainBusinessProfileId}</span>}
                </div>
              </div>

              {/* Compliance For — Pill Group */}
              <div className={styles.formRow}>
                <label className={styles.formLabel}>
                  Compliance For <span className={styles.requiredStar}>*</span>
                </label>
                <div className={styles.pillGroup}>
                  {COMPLIANCE_FOR_OPTIONS.map(opt => (
                    <button
                      key={opt.value}
                      type="button"
                      className={`${styles.pill} ${form.complianceFor === opt.value ? styles.pillActive : ''}`}
                      onClick={() => handleComplianceForChange(opt.value)}
                      id={`pill-${opt.value.toLowerCase()}`}
                    >
                      <span className={styles.pillIcon}>{opt.icon}</span>
                      {opt.value}
                    </button>
                  ))}
                </div>
                {errors.complianceFor && <span className={styles.errorMsg}>{errors.complianceFor}</span>}
              </div>

              {/* Dynamic Name Lookup */}
              {showDynamicField && (
                <div className={styles.formRow}>
                  <label className={styles.formLabel}>
                    {form.complianceFor} Name <span className={styles.requiredStar}>*</span>
                  </label>
                  <div className={errors.nameField ? styles.inputError : ''}>
                    <div className={styles.selectWrapper}>
                      <select
                        className={styles.select}
                        value={getDynamicValue() ?? ''}
                        onChange={e => setDynamicValue(e.target.value ? Number(e.target.value) : null)}
                        id={`select-${form.complianceFor.toLowerCase()}`}
                      >
                        <option value="">— Select {form.complianceFor} —</option>
                        {getDynamicOptions().map(opt => (
                          <option key={opt.key} value={opt.key}>{opt.text}</option>
                        ))}
                      </select>
                    </div>
                    {errors.nameField && <span className={styles.errorMsg}>{errors.nameField}</span>}
                  </div>
                </div>
              )}

              {/* Compliance Type */}
              <div className={styles.formRow}>
                <label className={styles.formLabel}>
                  Compliance Type <span className={styles.requiredStar}>*</span>
                </label>
                <div className={errors.complianceTypeId ? styles.inputError : ''}>
                  <div className={styles.selectWrapper}>
                    <select
                      className={styles.select}
                      value={form.complianceTypeId ?? ''}
                      onChange={e => set('complianceTypeId', e.target.value ? Number(e.target.value) : null)}
                      id="select-compliance-type"
                    >
                      <option value="">— Select Compliance Type —</option>
                      {lookups.complianceTypes.map(opt => (
                        <option key={opt.key} value={opt.key}>{opt.text}</option>
                      ))}
                    </select>
                  </div>
                  {errors.complianceTypeId && <span className={styles.errorMsg}>{errors.complianceTypeId}</span>}
                </div>
              </div>

            </div>
          </div>

          {/* ▸ Card 2: Booking / Document Details */}
          <div className={styles.card}>
            <div className={styles.cardHeader}>
              <div className={styles.cardHeaderAccent} />
              <span className={styles.cardHeaderTitle}>Compliance Details</span>
              <span className={styles.cardHeaderIcon}>📄</span>
            </div>
            <div className={styles.cardBody}>

              {/* Is Booking Toggle */}
              <div className={styles.formRow}>
                <label className={styles.formLabel}>Booking Record?</label>
                <label
                  className={`${styles.toggleWrapper} ${form.isBooking ? styles.toggleActive : ''}`}
                  htmlFor="chk-is-booking"
                >
                  <input
                    type="checkbox"
                    id="chk-is-booking"
                    className={styles.toggleInput}
                    checked={form.isBooking}
                    onChange={e => handleIsBookingChange(e.target.checked)}
                  />
                  <span className={styles.toggleTrack} />
                  <span className={styles.toggleLabel}>
                    {form.isBooking ? '✅ This is a booking record' : 'Mark as a booking record'}
                  </span>
                </label>
              </div>

              <div className={styles.sectionDivider} />

              {/* BOOKING FIELDS */}
              {showBookingFields && (
                <>
                  <div className={styles.sectionBadge}>📅 Booking Information</div>
                  <div className={styles.formRowInline}>
                    <div className={styles.formRow}>
                      <label className={styles.formLabel}>Booking Date</label>
                      <input
                        type="date"
                        value={form.bookingDate}
                        onChange={e => set('bookingDate', e.target.value)}
                        className={styles.dateInput}
                        id="input-booking-date"
                      />
                    </div>
                    <div className={styles.formRow}>
                      <label className={styles.formLabel}>Booked With</label>
                      <input
                        type="text"
                        value={form.bookedWith}
                        onChange={e => set('bookedWith', e.target.value)}
                        className={styles.textInput}
                        placeholder="e.g. Training provider"
                        id="input-booked-with"
                      />
                    </div>
                  </div>
                </>
              )}

              {/* DOCUMENT FIELDS */}
              {showDocFields && (
                <>
                  <div className={styles.sectionBadge}>🗂 Document Information</div>

                  {/* Document # */}
                  <div className={styles.formRow}>
                    <label className={styles.formLabel}>Document #</label>
                    <input
                      type="text"
                      value={form.documentNumber}
                      onChange={e => set('documentNumber', e.target.value)}
                      className={styles.textInput}
                      placeholder="e.g. DOC-2024-001"
                      id="input-document-number"
                    />
                  </div>

                  {/* Issuing Authority + Issue Date */}
                  <div className={styles.formRowInline}>
                    <div className={styles.formRow}>
                      <label className={styles.formLabel}>Issuing Authority</label>
                      <input
                        type="text"
                        value={form.issuingAuthority}
                        onChange={e => set('issuingAuthority', e.target.value)}
                        className={styles.textInput}
                        placeholder="e.g. Work Safe QLD"
                        id="input-issuing-authority"
                      />
                    </div>
                    <div className={styles.formRow}>
                      <label className={styles.formLabel}>Issue Date</label>
                      <input
                        type="date"
                        value={form.issueDate}
                        onChange={e => set('issueDate', e.target.value)}
                        className={styles.dateInput}
                        id="input-issue-date"
                      />
                    </div>
                  </div>

                  {/* Renewal + Expiry Date */}
                  <div className={styles.formRowInline}>
                    <div className={styles.formRow}>
                      <label className={styles.formLabel}>Renewal Required?</label>
                      <label
                        className={`${styles.toggleWrapper} ${form.renewalNotRequired ? styles.toggleActive : ''}`}
                        htmlFor="chk-renewal"
                      >
                        <input
                          type="checkbox"
                          id="chk-renewal"
                          className={styles.toggleInput}
                          checked={form.renewalNotRequired}
                          onChange={e => set('renewalNotRequired', e.target.checked)}
                        />
                        <span className={styles.toggleTrack} />
                        <span className={styles.toggleLabel}>
                          {form.renewalNotRequired ? '🔁 No renewal needed' : 'Renewal required'}
                        </span>
                      </label>
                    </div>
                    {showExpiryDate && (
                      <div className={styles.formRow}>
                        <label className={styles.formLabel}>Expiry Date</label>
                        <input
                          type="date"
                          value={form.expiryDate}
                          onChange={e => set('expiryDate', e.target.value)}
                          className={styles.dateInput}
                          id="input-expiry-date"
                        />
                      </div>
                    )}
                  </div>
                </>
              )}

              {/* Notes */}
              <div className={styles.sectionDivider} />
              <div className={styles.formRow}>
                <label className={styles.formLabel}>📝 Notes</label>
                <textarea
                  value={form.notes}
                  onChange={e => set('notes', e.target.value)}
                  className={styles.textarea}
                  rows={4}
                  placeholder="Add any additional notes or comments here…"
                  id="input-notes"
                />
              </div>

            </div>
          </div>

        </div>{/* /leftPanel */}

        {/* RIGHT — Attachments */}
        <div className={styles.rightPanel}>
          <AttachmentPanelInline
            attachments={attachments}
            siteUrl={siteUrl}
            loading={attachmentsLoading}
            onUpload={onUploadAttachment}
            onRemove={onRemoveAttachment}
          />
        </div>

      </div>{/* /formBody */}
    </div>
  );
};

export default ComplianceForm;
