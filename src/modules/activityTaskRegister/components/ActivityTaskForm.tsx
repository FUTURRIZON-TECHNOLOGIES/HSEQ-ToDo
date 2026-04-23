import * as React from 'react';
import {
    Stack,
    PrimaryButton,
    DefaultButton,
    IconButton,
    Spinner,
    SpinnerSize,
    Text,
    Dialog,
    DialogType,
    DialogFooter,
    Dropdown
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ActivityTaskService } from '../services/ActivityTaskService';
import { IActivityTaskItem, IActivityTaskHazard, ISPLookup } from '../models/IActivityTask';
import RiskAssessment from './RiskAssessment';
import HazardManager from './HazardManager';
import Timeline from './Timeline';
import styles from './ActivityTaskForm.module.scss';

interface IActivityTaskFormProps {
    item?: IActivityTaskItem;
    context: WebPartContext;
    onClose: () => void;
    onSave: (payload: Partial<IActivityTaskItem>, mode: 'stay' | 'close' | 'new') => Promise<void>;
    onRefresh: () => void;
}

const ActivityTaskForm: React.FC<IActivityTaskFormProps> = ({
    item,
    context,
    onClose,
    onSave,
    onRefresh
}) => {
    const [formData, setFormData] = React.useState<Partial<IActivityTaskItem>>(
        item ? { ...item } : { Active: true, HighRiskWork: false }
    );
    const [hazards, setHazards] = React.useState<IActivityTaskHazard[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [saving, setSaving] = React.useState(false);

    const [lookups, setLookups] = React.useState({
        activities:      [] as ISPLookup[],
        workZones:       [] as ISPLookup[],
        businessProfiles:[] as ISPLookup[],
        hazardTypes:     [] as ISPLookup[]   // For the dialog picker
    });
    const [choices, setChoices] = React.useState({
        consequences: [] as string[],
        likelihoods:  [] as string[]
    });

    // Dialog state
    const [isHazardDialogOpen, setIsHazardDialogOpen] = React.useState(false);
    const [pendingHazardId, setPendingHazardId] = React.useState<number | undefined>(undefined);

    const service = React.useMemo(() => new ActivityTaskService(context), [context]);

    // ── Load all reference data ───────────────────────────────────────────────
    React.useEffect(() => {
        const load = async () => {
            setLoading(true);
            try {
                const [acts, zones, profiles, hazardTypes, cons, likes] = await Promise.all([
                    service.getLookupOptions('Risk Register Activity Type', 'Name'),
                    service.getLookupOptions('Risk Register Work Zone Type', 'Name'),
                    service.getLookupOptions('Business Profiles', 'Business Profile'),
                    service.getLookupOptions('Activity & Task Hazard Type', 'Hazard Typd'),
                    service.getChoiceOptions('Consequence'),
                    service.getChoiceOptions('Likelihood')
                ]);
                setLookups({ activities: acts, workZones: zones, businessProfiles: profiles, hazardTypes });
                setChoices({ consequences: cons, likelihoods: likes });

                if (item?.Id) {
                    const h = await service.getHazards(item.Id);
                    setHazards(h);
                }
            } catch (err) {
                console.error("[ActivityTaskForm] Load failed:", err);
            } finally {
                setLoading(false);
            }
        };
        load();
    }, [service, item?.Id]);

    // Sync formData when item changes (edit vs new)
    React.useEffect(() => {
        setFormData(item ? { ...item } : { Active: true, HighRiskWork: false });
    }, [item]);

    const handleFieldChange = (field: string, value: any) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    const handleSave = async (mode: 'stay' | 'close' | 'new') => {
        setSaving(true);
        try {
            await onSave(formData, mode);
            if (mode === 'new') {
                setFormData({ Active: true, HighRiskWork: false });
                setHazards([]);
            }
        } finally {
            setSaving(false);
        }
    };

    // ── Add Hazard dialog ──────────────────────────────────────────────────────
    const handleAddHazard = async () => {
        if (!pendingHazardId) return;
        const hazardType = lookups.hazardTypes.find(h => h.Id === pendingHazardId);
        if (!hazardType) return;

        if (formData.Id) {
            // Saved record — persist to SP
            try {
                const newHazard = await service.addHazard(formData.Id, hazardType.Title);
                setHazards(prev => [...prev, newHazard]);
            } catch (e) {
                alert("Failed to save hazard.");
            }
        } else {
            // Unsaved record — add locally (will not persist until the form is saved first)
            setHazards(prev => [...prev, {
                Id: -(Date.now()),   // negative = local-only
                Title: hazardType.Title,
                ActivityTaskRegisterId: 0
            }]);
        }

        setIsHazardDialogOpen(false);
        setPendingHazardId(undefined);
    };

    const handleDeleteHazard = async (id: number) => {
        if (!confirm("Are you sure you want to delete this hazard?")) return;
        if (id > 0) {
            try {
                await service.deleteHazard(id);
            } catch (e) {
                alert("Failed to delete hazard from server.");
                return;
            }
        }
        setHazards(prev => prev.filter(h => h.Id !== id));
    };

    // ── Render ─────────────────────────────────────────────────────────────────
    if (loading) {
        return (
            <div className={styles.formContainer}>
                <Stack verticalAlign="center" horizontalAlign="center" styles={{ root: { height: '100%' } }}>
                    <Spinner size={SpinnerSize.large} label="Loading form data..." />
                </Stack>
            </div>
        );
    }

    return (
        <div className={styles.formContainer}>
            {/* Toolbar */}
            <div className={styles.toolbar}>
                <div className={styles.title}>
                    Master Activity & Task Register: {formData.Id ? `#${formData.Id}` : 'New Item'}
                </div>
                <div className={styles.actions}>
                    <PrimaryButton
                        iconProps={{ iconName: 'Save' }}
                        text={saving ? 'Saving…' : 'Save'}
                        disabled={saving}
                        onClick={() => handleSave('stay')}
                        className={styles.btnSave}
                    />
                    <DefaultButton
                        iconProps={{ iconName: 'SaveAndClose' }}
                        text="Save & Close"
                        disabled={saving}
                        onClick={() => handleSave('close')}
                        className={styles.btnSecondary}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Refresh' }}
                        title="Refresh"
                        onClick={onRefresh}
                        className={styles.btnSecondary}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        title="Close"
                        onClick={onClose}
                        className={styles.btnSecondary}
                    />
                </div>
            </div>

            {/* Body */}
            <div className={styles.content}>
                {/* Left — Risk Detail */}
                <div className={styles.leftColumn}>
                    <div className={styles.section}>
                        <Text
                            variant="mediumPlus"
                            styles={{ root: { fontWeight: 'bold', display: 'block', marginBottom: 16 } }}
                        >
                            RISK DETAIL &amp; ASSESSMENT
                        </Text>
                        <RiskAssessment
                            data={formData}
                            onChange={handleFieldChange}
                            lookups={{
                                activities:       lookups.activities,
                                workZones:        lookups.workZones,
                                businessProfiles: lookups.businessProfiles
                            }}
                            choices={choices}
                        />
                    </div>
                </div>

                {/* Right — Hazards + Timeline */}
                <div className={styles.rightColumn}>
                    <div className={styles.section}>
                        <HazardManager
                            hazards={hazards}
                            onAdd={() => setIsHazardDialogOpen(true)}
                            onDelete={handleDeleteHazard}
                            onEdit={() => { /* TODO */ }}
                            revisedAssessment={{
                                consequence: formData.RevisedConsequence,
                                likelihood:  formData.RevisedLikelihood,
                                ranking:     formData.RevisedRanking
                            }}
                            onRevisedChange={handleFieldChange}
                            choices={choices}
                        />
                    </div>

                    <div className={styles.section}>
                        <Timeline history={[]} />
                    </div>
                </div>
            </div>

            {/* Add Hazard Dialog */}
            <Dialog
                hidden={!isHazardDialogOpen}
                onDismiss={() => { setIsHazardDialogOpen(false); setPendingHazardId(undefined); }}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Add New Hazard',
                    subText: 'Select a hazard type to add to this task.'
                }}
                modalProps={{ isBlocking: true }}
                minWidth={440}
            >
                <Dropdown
                    label="Hazard Type"
                    placeholder="-- Select a Hazard --"
                    options={lookups.hazardTypes.map(h => ({ key: h.Id, text: h.Title }))}
                    selectedKey={pendingHazardId ?? null}
                    onChange={(_, opt) => setPendingHazardId(opt?.key as number)}
                    styles={{ root: { marginTop: 12 } }}
                />
                <DialogFooter>
                    <PrimaryButton
                        text="Add Hazard"
                        disabled={!pendingHazardId}
                        onClick={handleAddHazard}
                    />
                    <DefaultButton
                        text="Cancel"
                        onClick={() => { setIsHazardDialogOpen(false); setPendingHazardId(undefined); }}
                    />
                </DialogFooter>
            </Dialog>
        </div>
    );
};

export default ActivityTaskForm;
