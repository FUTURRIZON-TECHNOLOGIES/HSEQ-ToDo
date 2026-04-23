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
    onSave: (payload: any, mode: 'stay' | 'close' | 'new') => Promise<void>;
    onRefresh: () => void;
}

const ActivityTaskForm: React.FC<IActivityTaskFormProps> = ({
    item,
    context,
    onClose,
    onSave,
    onRefresh
}) => {
    const [formData, setFormData] = React.useState<Partial<IActivityTaskItem>>(item || {
        Active: true,
        HighRiskWork: false
    });
    const [hazards, setHazards] = React.useState<IActivityTaskHazard[]>([]);
    const [loading, setLoading] = React.useState(false);
    const [lookups, setLookups] = React.useState({
        activities: [] as ISPLookup[],
        workZones: [] as ISPLookup[],
        businessProfiles: [] as ISPLookup[],
        hazards: [] as ISPLookup[]
    });
    const [choices, setChoices] = React.useState({
        consequences: [] as string[],
        likelihoods: [] as string[]
    });
    const [isHazardDialogVisible, setIsHazardDialogVisible] = React.useState(false);
    const [selectedHazardId, setSelectedHazardId] = React.useState<number | undefined>(undefined);

    const service = React.useMemo(() => new ActivityTaskService(context), [context]);

    React.useEffect(() => {
        const loadData = async () => {
            setLoading(true);
            try {
                // Load lookups and choices
                const [acts, zones, profiles, hazards, cons, likes] = await Promise.all([
                    service.getLookupOptions('Risk Register Activity Type', 'Name'),
                    service.getLookupOptions('Risk Register Work Zone Type', 'Name'),
                    service.getLookupOptions('Business Profiles', 'Business Profile'),
                    service.getLookupOptions('Activity & Task Hazard Type', 'Hazard Typd'),
                    service.getChoiceOptions('Consequence'),
                    service.getChoiceOptions('Likelihood')
                ]);

                setLookups({ 
                    activities: acts, 
                    workZones: zones, 
                    businessProfiles: profiles,
                    hazards: hazards
                });
                setChoices({ consequences: cons, likelihoods: likes });

                if (item?.Id) {
                    const hazardData = await service.getHazards(item.Id);
                    setHazards(hazardData);
                }
            } catch (error) {
                console.error("Failed to load form data", error);
            } finally {
                setLoading(false);
            }
        };

        loadData();
    }, [service, item]);

    const handleFieldChange = (field: string, value: any) => {
        setFormData(prev => ({ ...prev, [field]: value }));
    };

    const handleSave = async (mode: 'stay' | 'close' | 'new') => {
        setLoading(true);
        try {
            await onSave(formData, mode);
        } finally {
            setLoading(false);
        }
    };

    if (loading && !formData.Id) {
        return (
            <div className={styles.formContainer}>
                <Stack verticalAlign="center" horizontalAlign="center" grow>
                    <Spinner size={SpinnerSize.large} label="Loading form..." />
                </Stack>
            </div>
        );
    }

    return (
        <div className={styles.formContainer}>
            <div className={styles.toolbar}>
                <div className={styles.title}>
                    Master Activity & Task Register: {formData.Id || 'New Item'}
                </div>
                <div className={styles.actions}>
                    <PrimaryButton 
                        iconProps={{ iconName: 'Save' }} 
                        text="Save" 
                        onClick={() => handleSave('stay')}
                        className={styles.btnSave}
                    />
                    <DefaultButton 
                        iconProps={{ iconName: 'SaveAndClose' }} 
                        text="Save & Close" 
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

            <div className={styles.content}>
                <div className={styles.leftColumn}>
                    <div className={styles.section}>
                        <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold', display: 'block', marginBottom: 15 } }}>RISK DETAIL & ASSESSMENT</Text>
                        <RiskAssessment 
                            data={formData} 
                            onChange={handleFieldChange}
                            lookups={lookups}
                            choices={choices}
                        />
                    </div>
                </div>

                <div className={styles.rightColumn}>
                    <div className={styles.section}>
                        <HazardManager 
                            hazards={hazards}
                            onAdd={() => setIsHazardDialogVisible(true)}
                            onDelete={async (id) => {
                                if (confirm("Delete this hazard?")) {
                                    await service.deleteHazard(id);
                                    setHazards(prev => prev.filter(h => h.Id !== id));
                                }
                            }}
                            onEdit={() => {}}
                            revisedAssessment={{
                                consequence: formData.RevisedConsequence,
                                likelihood: formData.RevisedLikelihood,
                                ranking: formData.RevisedRanking
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

            <Dialog
                hidden={!isHazardDialogVisible}
                onDismiss={() => setIsHazardDialogVisible(false)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Add New Hazard',
                    subText: 'Select a hazard from the list below to add it to this task.'
                }}
                modalProps={{ isBlocking: false }}
            >
                <Dropdown
                    label="Select Hazard"
                    options={lookups.hazards.map(h => ({ key: h.Id, text: h.Title }))}
                    selectedKey={selectedHazardId}
                    onChange={(_, opt) => setSelectedHazardId(opt?.key as number)}
                />
                <DialogFooter>
                    <PrimaryButton 
                        text="Add" 
                        disabled={!selectedHazardId} 
                        onClick={async () => {
                            if (selectedHazardId && formData.Id) {
                                const hazard = lookups.hazards.find(h => h.Id === selectedHazardId);
                                if (hazard) {
                                    const newHazard = await service.addHazard(formData.Id, hazard.Id, hazard.Title);
                                    setHazards(prev => [...prev, newHazard]);
                                }
                            }
                            setIsHazardDialogVisible(false);
                            setSelectedHazardId(undefined);
                        }} 
                    />
                    <DefaultButton text="Cancel" onClick={() => setIsHazardDialogVisible(false)} />
                </DialogFooter>
            </Dialog>
        </div>
    );
};


export default ActivityTaskForm;
