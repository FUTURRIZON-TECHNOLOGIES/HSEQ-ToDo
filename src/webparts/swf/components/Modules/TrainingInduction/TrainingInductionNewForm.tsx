import * as React from 'react';
import { 
    TextField,
    IDropdownOption, 
    ComboBox,
    DatePicker, 
    Checkbox,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType
} from '@fluentui/react';
import styles from './TrainingInduction.module.scss';
import { PrimaryButton, DefaultButton, Icon } from '@fluentui/react';
import { ITrainingInductionItem } from '../../../models/ITrainingInductionItem';
import { SPService } from '../../../services/SPService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITrainingInductionNewFormProps {
    spService: SPService;
    context: WebPartContext;
    onClose: () => void;
    onSave: (payload: any, mode: 'stay' | 'close' | 'new') => Promise<void>;
    preLoadedLookups?: { participants: IDropdownOption[], businessProfiles: IDropdownOption[], employees: IDropdownOption[], trainingTypes: IDropdownOption[] };
}


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

const TrainingInductionNewForm: React.FC<ITrainingInductionNewFormProps> = (props) => {
    const [formData, setFormData] = React.useState<Partial<ITrainingInductionItem>>({
        Title: '',
        TrainingFor: '',
        SendInvitation: true
    });
    const [saving, setSaving] = React.useState(false);
    const [error, setError] = React.useState<string | null>(null);
    const [validationErrors, setValidationErrors] = React.useState<Record<string, string>>({});
    const [lookupOptions, setLookupOptions] = React.useState<{ 
        participants: IDropdownOption[], 
        businessProfiles: IDropdownOption[],
        employees: IDropdownOption[],
        trainingTypes: IDropdownOption[]
    }>({
        participants: props.preLoadedLookups?.participants || [],
        businessProfiles: props.preLoadedLookups?.businessProfiles || [],
        employees: props.preLoadedLookups?.employees || [],
        trainingTypes: props.preLoadedLookups?.trainingTypes || []
    });

    React.useEffect(() => {
        if (!props.preLoadedLookups) {
            const load = async () => {
                try {
                    const [pts, bps, emps, ttypes] = await Promise.all([
                        props.spService.getLookupOptions('Contacts', 'Employee Name'),
                        props.spService.getLookupOptions('BusinessProfiles', 'BusinessProfile'),
                        props.spService.getLookupOptions('Employees', 'Employee Name'),
                        props.spService.getLookupOptions('TrainingTypes', 'Title')
                    ]);
                    setLookupOptions({
                        participants: pts.map(p => ({ key: p.Id, text: p.Title })),
                        businessProfiles: bps.map(b => ({ key: b.Id, text: b.Title })),
                        employees: emps.map(e => ({ key: e.Id, text: e.Title })),
                        trainingTypes: ttypes.map(t => ({ key: t.Title, text: t.Title }))
                    });
                } catch (e) { console.error('Lookup load failed', e); }
            };
            load();
        } else {
            setLookupOptions({ ...props.preLoadedLookups });
        }
    }, [props.preLoadedLookups]);

    const validate = (): boolean => {
        const newErrors: Record<string, string> = {};
        if (!formData.Type) newErrors.Type = 'System Form is required.';
        if (!formData.Participant?.Id) newErrors.Participant = 'Participants is required.';
        if (!formData.ScheduledDate) newErrors.ScheduledDate = 'Schedule Date is required.';
        
        setValidationErrors(newErrors);
        return Object.keys(newErrors).length === 0;
    };

    const handleFieldChange = (field: keyof ITrainingInductionItem, value: any) => {
        setFormData(prev => ({ ...prev, [field]: value }));
        if (validationErrors[field]) {
            setValidationErrors(prev => {
                const updated = { ...prev };
                delete updated[field];
                return updated;
            });
        }
    };

    const handleSave = async (mode: 'stay' | 'close' | 'new') => {
        if (!validate()) return;
        setSaving(true);
        setError(null);
        try {
            const payload: any = { 
                Title: formData.Type || '', 
                Type: formData.Type || '', 
                TrainingType: formData.Type || '',
                "System Form": formData.Type || '',
                "SystemForm": formData.Type || '',
                TrainingFor: formData.TrainingFor || '',
                SendInvitation: formData.SendInvitation,
                Status: 'Scheduled'
            };

            // Business Profile
            if (formData.BusinessProfile?.Id) {
                payload.BusinessProfileId = formData.BusinessProfile.Id;
                payload["Business ProfileId"] = formData.BusinessProfile.Id;
                payload["Business_x0020_ProfileId"] = formData.BusinessProfile.Id;
            }

            // Lookups / People
            if (formData.Participant?.Id) {
                payload.ParticipantId = formData.Participant.Id;
                payload.ParticipantsId = formData.Participant.Id;
            }
            if (formData.Manager?.Id) payload.ManagerId = formData.Manager.Id;
            if (formData.Supervisors?.Id) payload.SupervisorsId = formData.Supervisors.Id;
            if (formData.Coordinator?.Id) payload.CoordinatorId = formData.Coordinator.Id;
            if (formData.ScheduledDate) payload.ScheduledDate = formData.ScheduledDate;
            
            // Hardened mapping for Send Invitation variants
            payload.SendInvitations = formData.SendInvitation;
            payload.Send_x0020_Invitation = formData.SendInvitation;
            payload.Send_x0020_Invitations = formData.SendInvitation;

            await props.onSave(payload, mode);
            
            if (mode === 'new') {
                setFormData({ 
                    Type: '', 
                    TrainingFor: '', 
                    SendInvitation: true,
                    Participant: undefined,
                    BusinessProfile: undefined,
                    Manager: undefined,
                    Supervisors: undefined,
                    Coordinator: undefined,
                    ScheduledDate: undefined
                });
                setValidationErrors({});
            }
        } catch (e) {
            console.error('Save error', e);
            setError(`Save failed: ${e.message || 'Check console for details.'}`);
        } finally {
            setSaving(false);
        }
    };

    const fieldLabelStyle: React.CSSProperties = { 
        width: 140, 
        minWidth: 140, 
        fontSize: 13, 
        color: '#323130', 
        fontWeight: 600,
        paddingTop: 6
    };
    const fieldRowStyle: React.CSSProperties = { 
        display: 'flex', 
        alignItems: 'flex-start', 
        marginBottom: 15, 
        gap: 10 
    };
    const inputContainerStyle: React.CSSProperties = {
        flex: 1,
        maxWidth: 400
    };

    return (
        <div className={styles.todoForm}>
            {/* Header / Toolbar (Matching ToDo Style) */}
            <div className={styles.toolbar}>
                <div className={styles.formTitle}>
                    <Icon iconName="Education" className={styles.formTitleIcon} />
                    <span>Training: New — New Activity</span>
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
                        onClick={() => window.location.reload()}
                    />
                    <DefaultButton
                        className={`${styles.btnAction} ${styles.btnClose}`}
                        iconProps={{ iconName: 'Cancel' }}
                        text="Close"
                        onClick={props.onClose}
                    />
                </div>
            </div>

            {error && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { margin: '10px 30px' } }}>{error}</MessageBar>}

            <div style={{ padding: '30px', maxHeight: '70vh', overflowY: 'auto' }}>
                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>System Form</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="(required)" 
                            options={lookupOptions.trainingTypes.length > 0 ? lookupOptions.trainingTypes : TYPE_OPTIONS} 
                            selectedKey={formData.Type}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => {
                                handleFieldChange('Type', opt?.key);
                                if (opt) handleFieldChange('Title', opt.text);
                            }}
                            errorMessage={validationErrors.Type}
                        />
                    </div>
                </div>



                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Training For</span>
                    <div style={inputContainerStyle}>
                        <TextField 
                            value={formData.TrainingFor || ''} 
                            onChange={(_, val) => handleFieldChange('TrainingFor', val)}
                        />
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Business Profile</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="Select Business Profile"
                            options={lookupOptions.businessProfiles} 
                            selectedKey={formData.BusinessProfile?.Id}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => handleFieldChange('BusinessProfile', { Id: opt?.key, Title: opt?.text })}
                        />
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Participants</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="(required)" 
                            options={lookupOptions.participants} 
                            selectedKey={formData.Participant?.Id}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => handleFieldChange('Participant', { Id: opt?.key, Title: opt?.text })}
                            errorMessage={validationErrors.Participant}
                        />
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Schedule Date</span>
                    <div style={inputContainerStyle}>
                        <DatePicker 
                            placeholder="(required)" 
                            value={formData.ScheduledDate ? new Date(formData.ScheduledDate) : undefined}
                            onSelectDate={(date) => handleFieldChange('ScheduledDate', date?.toISOString())}
                            // DatePicker doesn't have errorMessage, so we use a sub-text or style
                        />
                        {validationErrors.ScheduledDate && <span style={{ color: '#a4262c', fontSize: 12, marginTop: 4 }}>{validationErrors.ScheduledDate}</span>}
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Manager</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="Select Manager"
                            options={lookupOptions.employees} 
                            selectedKey={formData.Manager?.Id}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => handleFieldChange('Manager', { Id: opt?.key, Title: opt?.text })}
                        />
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Supervisor</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="Select Supervisor"
                            options={lookupOptions.employees} 
                            selectedKey={formData.Supervisors?.Id}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => handleFieldChange('Supervisors', { Id: opt?.key, Title: opt?.text })}
                        />
                    </div>
                </div>

                <div style={fieldRowStyle}>
                    <span style={fieldLabelStyle}>Coordinator</span>
                    <div style={inputContainerStyle}>
                        <ComboBox 
                            placeholder="Select Coordinator"
                            options={lookupOptions.employees} 
                            selectedKey={formData.Coordinator?.Id}
                            allowFreeform={false}
                            autoComplete='on'
                            onChange={(_, opt) => handleFieldChange('Coordinator', { Id: opt?.key, Title: opt?.text })}
                        />
                    </div>
                </div>

                <div style={{ ...fieldRowStyle, marginTop: 10 }}>
                    <span style={fieldLabelStyle}>Send Invitations</span>
                    <Checkbox checked={formData.SendInvitation} onChange={(_, checked) => handleFieldChange('SendInvitation', !!checked)} />
                </div>
            </div>

            {saving && (
                <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(255,255,255,0.7)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
                    <Spinner size={SpinnerSize.large} label="Saving Record..." />
                </div>
            )}
        </div>
    );
};

export default TrainingInductionNewForm;
