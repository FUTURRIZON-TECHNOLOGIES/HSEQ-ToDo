import * as React from 'react';
import {
    TextField,
    DatePicker,
    Dropdown,
    IDropdownOption,
    Checkbox,
    DefaultButton,
    PrimaryButton,
    Icon,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType
} from '@fluentui/react';
import { IToDoItem } from '../../../models/IToDoItem';
import { SPService } from '../../../services/SPService';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import AttachmentControl from '../../Shared/AttachmentControl';
import styles from './ToDoForm.module.scss';
import { REGARDING_DYNAMIC_FIELDS, getRegardingFieldLabel } from './RegardingConfig';

export interface IToDoFormProps {
    item: IToDoItem | null;
    spService: SPService;
    context: WebPartContext;
    onSave: (payload: any, mode?: 'stay' | 'close' | 'new') => void;
    onClose: () => void;
}

// ── Collapsible Section ───────────────────────────────────────────────────────
const Section: React.FC<{ title: string; icon: string; defaultOpen?: boolean; children: React.ReactNode }> = ({
    title, icon, defaultOpen = true, children
}) => {
    const [open, setOpen] = React.useState(defaultOpen);
    return (
        <div className={styles.section}>
            <div className={styles.sectionHeader} onClick={() => setOpen(!open)}>
                <Icon iconName={icon} className={styles.sectionIcon} />
                <span className={styles.sectionTitle}>{title}</span>
                <Icon iconName={open ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
            </div>
            {open && <div className={styles.sectionBody}>{children}</div>}
        </div>
    );
};

// ── Helper: build initial formData from item (ensures IDs are populated) ─────
const buildFormData = (item: IToDoItem | null): IToDoItem => {
    if (!item) return { Title: '', CompletedPercent: 0 };
    return {
        ...item,
        // Ensure dropdown selectedKey values are populated from lookup objects
        StatusId:         item.StatusId         != null ? item.StatusId         : item.Status?.Id,
        CategoryId:       item.CategoryId       != null ? item.CategoryId       : item.Category?.Id,
        ClassificationId: item.ClassificationId != null ? item.ClassificationId : item.Classification?.Id,
        PriorityId:       item.PriorityId       != null ? item.PriorityId       : item.Priority?.Id,
    };
};

// ── Main Form ─────────────────────────────────────────────────────────────────
const ToDoForm: React.FC<IToDoFormProps> = ({ item, spService, context, onSave, onClose }) => {
    const isNew = !item?.Id;

    const [formData, setFormData] = React.useState<IToDoItem>(buildFormData(item));
    const [attachments, setAttachments] = React.useState<any[]>([]);
    const [pendingAttachments, setPendingAttachments] = React.useState<File[]>([]);
    const [options, setOptions] = React.useState<{
        status: IDropdownOption[];
        category: IDropdownOption[];
        classification: IDropdownOption[];
        priority: IDropdownOption[];
        regarding: IDropdownOption[];
    }>({ status: [], category: [], classification: [], priority: [], regarding: [] });
    const [loadingOptions, setLoadingOptions] = React.useState(true);
    const [saving, setSaving] = React.useState(false);
    const [error, setError] = React.useState<string | null>(null);

    // ── Dynamic Field State for 'Regarding' ──────────────
    const [dynamicFieldKey, setDynamicFieldKey] = React.useState<string>("");
    const [dynamicFieldValue, setDynamicFieldValue] = React.useState<string>("");

    const regardingChoices: IDropdownOption[] = [
        'Audit & Inspection', 'Clients', 'Compliance Register', 'Employee', 'Incident',
        'Leads', 'Meetings', 'Project', 'Proposal', 'Subcontractor',
        'Subcontractor Employee', 'Submission', 'Training & Induction', 'Vehicle & Plant'
    ].map(c => ({ key: c, text: c }));

    // ── Fetch attachments ─────────────────────────────────────────────────────
    const fetchAttachments = React.useCallback(async () => {
        if (item?.Id) {
            try { setAttachments(await spService.getAttachments(item.Id)); }
            catch (e) { console.error('Attachment fetch failed', e); }
        }
    }, [item?.Id]);

    // ── Load options + reset form when item changes ───────────────────────────
    React.useEffect(() => {
        // ───────────────────────────────────────────────────────────────────────
        // CRITICAL: reset formData including IDs from lookup objects
        // ───────────────────────────────────────────────────────────────────────
        setFormData(buildFormData(item));
        setPendingAttachments([]);

        // ── Resolve and Set Dynamic Field for Edit mode ──
        if (item?.Regarding) {
            const fieldKey = REGARDING_DYNAMIC_FIELDS[item.Regarding];
            setDynamicFieldKey(fieldKey || "");
            setDynamicFieldValue(fieldKey ? (item as any)[fieldKey] || "" : "");
        } else {
            setDynamicFieldKey("");
            setDynamicFieldValue("");
        }

        const loadOptions = async () => {
            setLoadingOptions(true);
            try {
                const [statusOpts, categoryOpts, classificationOpts, priorityOpts] = await Promise.all([
                    spService.getLookupOptions('TaskStatus', 'Name'),
                    spService.getLookupOptions('TaskCategory', 'Name'),
                    spService.getLookupOptions('ClassificationType', 'Name'),
                    spService.getLookupOptions('TaskPriorities', 'Name'),
                ]);
                setOptions({
                    status:         statusOpts.map(o => ({ key: o.Id, text: o.Title })),
                    category:       categoryOpts.map(o => ({ key: o.Id, text: o.Title })),
                    classification: classificationOpts.map(o => ({ key: o.Id, text: o.Title })),
                    priority:       priorityOpts.map(o => ({ key: o.Id, text: o.Title })),
                    regarding:      regardingChoices,
                });
            } catch (e) {
                console.error('Dropdown load failed', e);
            } finally {
                setLoadingOptions(false);
            }
        };
        loadOptions();
        fetchAttachments();
    }, [item?.Id]);

    // ── Save handler ──────────────────────────────────────────────────────────
    const handleSaveInternal = async (mode: 'stay' | 'close' | 'new') => {
        const requiredFields = [
            { key: 'Title', label: 'Subject' },
            { key: 'Regarding', label: 'Regarding' },
            { key: 'CategoryId', label: 'Category' },
            { key: 'ClassificationId', label: 'Classification' },
            { key: 'PriorityId', label: 'Priority' },
            { key: 'StatusId', label: 'Status' }
        ];

        for (const field of requiredFields) {
            const val = (formData as any)[field.key];
            if (val === undefined || val === null || (typeof val === 'string' && !val.trim())) {
                setError(`${field.label} is required.`);
                return;
            }
        }

        setError(null);
        setSaving(true);
        const getInternal = (displayName: string, fallback: string) =>
            (spService as any).getInternalName(displayName, fallback);

        const payload: any = {
            Description:        formData.Description,
            Regarding:          formData.Regarding,
            DueDate:            formData.DueDate,
            StartDate:          formData.StartDate,
            Resolution:         formData.Resolution,
            EstimationDuration: formData.EstimationDuration,
            ActualDuration:     formData.ActualDuration,
        };
        const subjectInternal = getInternal('Subject', 'Title');
        payload[subjectInternal] = formData.Title;
        if (subjectInternal !== 'Title') payload['Title'] = formData.Title;

        payload[`${getInternal('Status',         'Status')}Id`]           = formData.StatusId;
        payload[`${getInternal('Category',       'Category')}Id`]         = formData.CategoryId;
        payload[`${getInternal('Classification', 'Classification')}Id`]   = formData.ClassificationId;
        payload[`${getInternal('Priority',       'Priority')}Id`]         = formData.PriorityId;
        payload[`${getInternal('Task Owner',     'TaskOwner')}Id`]        = formData.TaskOwnerId;
        payload[`${getInternal('Assigne Internal','AssigneInternal')}Id`] = formData.AssigneeInternalId;
        payload[`${getInternal('Assigne External','AssigneExternal')}Id`] = formData.AssigneeExternalId;
        payload[getInternal('Completed %',       'CompletedPercent')]     = formData.CompletedPercent;
        payload[getInternal('Email Notification','EmailNotifications')]   = formData.EmailNotifications;
        payload[getInternal('Completion Date',   'CompletionDate')]       = formData.CompletionDate;

        // ── Dynamic 'Regarding' Fields Integration ────────
        if (dynamicFieldKey && dynamicFieldValue) {
            payload[dynamicFieldKey] = dynamicFieldValue;
        }

        payload.PendingAttachments = pendingAttachments;

        try { await onSave(payload, mode); }
        finally { setSaving(false); }
    };

    // ── Attachment handlers ───────────────────────────────────────────────────
    const handleUpload = async (file: File) => {
        if (item?.Id) {
            await spService.uploadAttachment(item.Id, file);
            await fetchAttachments();
        } else {
            setPendingAttachments(prev => [...prev, file]);
        }
    };
    const handleDeleteAttachment = async (fileName: string) => {
        if (item?.Id) {
            if (confirm(`Delete "${fileName}"?`)) {
                await spService.deleteAttachment(item.Id, fileName);
                await fetchAttachments();
            }
        } else {
            setPendingAttachments(prev => prev.filter(f => f.name !== fileName));
        }
    };

    const allAttachments = React.useMemo(() => {
        const existing = attachments.map(a => ({ ...a, isPending: false }));
        const pending = pendingAttachments.map(f => ({
            FileName: f.name,
            ServerRelativeUrl: '#',
            isPending: true
        }));
        return [...existing, ...pending];
    }, [attachments, pendingAttachments]);

    // ── People picker shared props ────────────────────────────────────────────
    const pickerProps = {
        context: context as any,
        webAbsoluteUrl: context.pageContext.web.absoluteUrl,
        principalTypes: [PrincipalType.User],
        resolveDelay: 1000,
        ensureUser: true,
    };

    // ── Metadata display ──────────────────────────────────────────────────────
    const fmtMeta = (type: 'Author' | 'Created' | 'Editor' | 'Modified') => {
        if (!item) return 'Auto-populated on save';
        if (type === 'Author')   return item.Author?.Title  || '—';
        if (type === 'Created')  return item.Created  ? new Date(item.Created).toLocaleString()  : '—';
        if (type === 'Editor')   return item.Editor?.Title  || '—';
        if (type === 'Modified') return item.Modified ? new Date(item.Modified).toLocaleString() : '—';
        return '—';
    };

    if (loadingOptions) {
        return (
            <div className={styles.loadingWrapper}>
                <Spinner size={SpinnerSize.large} label="Loading form data…" />
            </div>
        );
    }

    return (
        <div className={styles.todoForm}>

            {/* ── Toolbar ── */}
            <div className={styles.toolbar}>
                <div className={styles.formTitle}>
                    <Icon iconName="TaskManager" className={styles.formTitleIcon} />
                    <span>{isNew ? 'Task: New — New Activity' : `Task: ${item!.Id} — ${item!.Title}`}</span>
                </div>
                <div className={styles.toolbarActions}>
                    <PrimaryButton
                        className={styles.btnSave}
                        iconProps={{ iconName: 'Save' }}
                        text={saving ? 'Saving…' : 'Save'}
                        disabled={saving}
                        onClick={() => handleSaveInternal('stay')}
                    />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'SaveAs' }}
                        text="Save & New"
                        disabled={saving}
                        onClick={() => handleSaveInternal('new')}
                    />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'SaveAndClose' }}
                        text="Save & Close"
                        disabled={saving}
                        onClick={() => handleSaveInternal('close')}
                    />
                    <div className={styles.toolbarDivider} />
                    <DefaultButton
                        className={styles.btnAction}
                        iconProps={{ iconName: 'Refresh' }}
                        text="Refresh"
                        disabled={saving}
                        onClick={() => fetchAttachments()}
                    />
                    <DefaultButton
                        className={`${styles.btnAction} ${styles.btnClose}`}
                        iconProps={{ iconName: 'Cancel' }}
                        text="Close"
                        onClick={onClose}
                    />
                </div>
            </div>

            {error && (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    onDismiss={() => setError(null)}
                    className={styles.errorBar}
                >
                    {error}
                </MessageBar>
            )}

            {/* ── Body ── */}
            <div className={styles.scrollContent}>

                {/* ── Left column ── */}
                <div className={styles.leftColumn}>

                    {/* General Information */}
                    <Section title="GENERAL INFORMATION" icon="Info">
                        <div className={styles.fieldRow}>
                            <div className={styles.fieldFull}>
                                <TextField
                                    label="Subject" required
                                    value={formData.Title}
                                    onChange={(_, val) => setFormData({ ...formData, Title: val || '' })}
                                    placeholder="Enter task subject"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldFull}>
                                <TextField
                                    label="Description" multiline rows={3}
                                    value={formData.Description}
                                    onChange={(_, val) => setFormData({ ...formData, Description: val || '' })}
                                    placeholder="Task description…"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <Dropdown
                                    label="Regarding" required
                                    options={options.regarding}
                                    selectedKey={formData.Regarding}
                                    onChange={(_, opt) => {
                                        const newVal = opt?.key as string;
                                        setFormData({ ...formData, Regarding: newVal });
                                        
                                        // Set correct internal field key using mapping and reset previous value
                                        const fieldKey = REGARDING_DYNAMIC_FIELDS[newVal];
                                        setDynamicFieldKey(fieldKey || "");
                                        setDynamicFieldValue("");
                                    }}
                                    placeholder="Select…"
                                />
                            </div>
                            <div className={styles.fieldHalf}>
                                {dynamicFieldKey && (
                                    <TextField
                                        label={getRegardingFieldLabel(dynamicFieldKey)}
                                        value={dynamicFieldValue}
                                        onChange={(_, val) => setDynamicFieldValue(val || '')}
                                        placeholder={`Enter ${getRegardingFieldLabel(dynamicFieldKey).toLowerCase()}…`}
                                    />
                                )}
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <Dropdown
                                    label="Category" required
                                    options={options.category}
                                    selectedKey={formData.CategoryId}
                                    onChange={(_, opt) => setFormData({ ...formData, CategoryId: opt?.key as number })}
                                    placeholder="Select category…"
                                />
                            </div>
                            <div className={styles.fieldHalf}>
                                {/* Classification lookup from ClassificationType list */}
                                <Dropdown
                                    label="Classification" required
                                    options={options.classification}
                                    selectedKey={formData.ClassificationId}
                                    onChange={(_, opt) => setFormData({ ...formData, ClassificationId: opt?.key as number })}
                                    placeholder="Select classification…"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <PeoplePicker
                                    {...pickerProps}
                                    titleText="Task Owner"
                                    personSelectionLimit={1}
                                    defaultSelectedUsers={formData.TaskOwner?.EMail ? [formData.TaskOwner.EMail] : []}
                                    onChange={people => setFormData({
                                        ...formData,
                                        TaskOwnerId: people.length > 0 ? (people[0] as any).id : undefined
                                    })}
                                />
                            </div>
                            <div className={styles.fieldHalf}>
                                <PeoplePicker
                                    {...pickerProps}
                                    titleText="Assignee (Internal)"
                                    personSelectionLimit={1}
                                    defaultSelectedUsers={formData.AssigneeInternal?.EMail ? [formData.AssigneeInternal.EMail] : []}
                                    onChange={people => setFormData({
                                        ...formData,
                                        AssigneeInternalId: people.length > 0 ? (people[0] as any).id : undefined
                                    })}
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <PeoplePicker
                                    {...pickerProps}
                                    titleText="Assignee (External)"
                                    personSelectionLimit={1}
                                    defaultSelectedUsers={formData.AssigneeExternal?.EMail ? [formData.AssigneeExternal.EMail] : []}
                                    onChange={people => setFormData({
                                        ...formData,
                                        AssigneeExternalId: people.length > 0 ? (people[0] as any).id : undefined
                                    })}
                                />
                            </div>
                            <div className={styles.fieldHalf} />
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <DatePicker
                                    label="Due Date"
                                    value={formData.DueDate ? new Date(formData.DueDate) : undefined}
                                    onSelectDate={date => setFormData({ ...formData, DueDate: date?.toISOString() })}
                                    placeholder="Select date…"
                                />
                            </div>
                            <div className={styles.fieldHalf}>
                                <Dropdown
                                    label="Priority" required
                                    options={options.priority}
                                    selectedKey={formData.PriorityId}
                                    onChange={(_, opt) => setFormData({ ...formData, PriorityId: opt?.key as number })}
                                    placeholder="Select priority…"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldThird}>
                                <Dropdown
                                    label="Status" required
                                    options={options.status}
                                    selectedKey={formData.StatusId}
                                    onChange={(_, opt) => setFormData({ ...formData, StatusId: opt?.key as number })}
                                    placeholder="Select status…"
                                />
                            </div>
                            <div className={`${styles.fieldThird} ${styles.checkboxField}`}>
                                <label className={styles.checkboxLabel}>Email Notification</label>
                                <Checkbox
                                    checked={formData.EmailNotifications}
                                    onChange={(_, val) => setFormData({ ...formData, EmailNotifications: val })}
                                />
                            </div>
                            <div className={styles.fieldThird} />
                        </div>
                    </Section>

                    {/* Internal Section */}
                    <Section title="INTERNAL" icon="Settings">
                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <DatePicker
                                    label="Start Date"
                                    value={formData.StartDate ? new Date(formData.StartDate) : undefined}
                                    onSelectDate={date => setFormData({ ...formData, StartDate: date?.toISOString() })}
                                    placeholder="Select date…"
                                />
                            </div>
                            <div className={styles.fieldHalf}>
                                <DatePicker
                                    label="Completion Date"
                                    value={formData.CompletionDate ? new Date(formData.CompletionDate) : undefined}
                                    onSelectDate={date => setFormData({ ...formData, CompletionDate: date?.toISOString() })}
                                    placeholder="Select date…"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                <TextField
                                    label="% Completed"
                                    type="number"
                                    min={0} max={100}
                                    value={String(formData.CompletedPercent ?? 0)}
                                    onChange={(_, val) => setFormData({ ...formData, CompletedPercent: Number(val) })}
                                />
                                <div className={styles.progressBar}>
                                    <div
                                        className={styles.progressFill}
                                        style={{ width: `${Math.min(formData.CompletedPercent ?? 0, 100)}%` }}
                                    />
                                </div>
                            </div>
                            <div className={styles.fieldHalf}>
                                {/* Est. Duration as DatePicker per user request */}
                                <DatePicker
                                    label="Est. Duration Date"
                                    value={formData.EstimationDuration ? new Date(formData.EstimationDuration) : undefined}
                                    onSelectDate={date => setFormData({ ...formData, EstimationDuration: date?.toISOString() })}
                                    placeholder="Select date…"
                                />
                            </div>
                        </div>

                        <div className={styles.fieldRow}>
                            <div className={styles.fieldHalf}>
                                {/* Actual Duration as DatePicker per user request */}
                                <DatePicker
                                    label="Actual Duration Date"
                                    value={formData.ActualDuration ? new Date(formData.ActualDuration) : undefined}
                                    onSelectDate={date => setFormData({ ...formData, ActualDuration: date?.toISOString() })}
                                    placeholder="Select date…"
                                />
                            </div>
                            <div className={styles.fieldHalf} />
                        </div>

                        {/* Metadata */}
                        <div className={styles.metaGrid}>
                            <div className={styles.metaItem}>
                                <span className={styles.metaLabel}>Created By</span>
                                <span className={styles.metaValue}>{fmtMeta('Author')}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <span className={styles.metaLabel}>Created On</span>
                                <span className={styles.metaValue}>{fmtMeta('Created')}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <span className={styles.metaLabel}>Updated By</span>
                                <span className={styles.metaValue}>{fmtMeta('Editor')}</span>
                            </div>
                            <div className={styles.metaItem}>
                                <span className={styles.metaLabel}>Updated On</span>
                                <span className={styles.metaValue}>{fmtMeta('Modified')}</span>
                            </div>
                        </div>
                    </Section>

                    {/* Attachments */}
                    <Section title="ATTACHMENTS" icon="Attach">
                        <AttachmentControl
                            attachments={allAttachments as any}
                            onUpload={handleUpload}
                            onDelete={handleDeleteAttachment}
                        />
                    </Section>
                </div>

                {/* ── Right column ── */}
                <div className={styles.rightColumn}>

                    {/* Resolution — rich text with visible toolbar */}
                    <Section title="RESOLUTION" icon="QuickNote">
                        <div className={styles.richTextWrapper}>
                            <label className="ms-Label" style={{ fontWeight: 600, fontSize: 12, marginBottom: 8, display: 'block' }}>
                                Resolution / Final Notes
                            </label>
                            <RichText
                                value={formData.Resolution || ''}
                                onChange={text => { setFormData({ ...formData, Resolution: text }); return text; }}
                                isEditMode={true}
                            />
                        </div>
                    </Section>

                    {/* Timeline */}
                    <Section title="TIMELINE" icon="Timeline">
                        <div className={styles.timeline}>
                            {item?.Modified && (
                                <div className={styles.timelineEvent}>
                                    <div className={styles.timelineDot} />
                                    <div className={styles.timelineContent}>
                                        <strong>{item.Editor?.Title || 'User'}</strong>
                                        <span className={styles.timelineDate}>
                                            {new Date(item.Modified).toLocaleString()}
                                        </span>
                                        <span className={styles.timelineDesc}>Task Detail Updated</span>
                                    </div>
                                </div>
                            )}
                            {item?.Created && (
                                <div className={styles.timelineEvent}>
                                    <div className={`${styles.timelineDot} ${styles.timelineDotLast}`} />
                                    <div className={styles.timelineContent}>
                                        <strong>{item.Author?.Title || 'User'}</strong>
                                        <span className={styles.timelineDate}>
                                            {new Date(item.Created).toLocaleString()}
                                        </span>
                                        <span className={styles.timelineDesc}>New Task Created</span>
                                    </div>
                                </div>
                            )}
                            {!item?.Created && (
                                <p className={styles.timelineEmpty}>Timeline will populate after saving.</p>
                            )}
                        </div>
                    </Section>
                </div>
            </div>
        </div>
    );
};

export default ToDoForm;
