import * as React from 'react';
import { 
    Stack, 
    PrimaryButton, 
    DefaultButton, 
    IconButton,
    Spinner,
    SpinnerSize,
    Text
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
        businessProfiles: [] as ISPLookup[]
    });
    const [choices, setChoices] = React.useState({
        consequences: [] as string[],
        likelihoods: [] as string[]
    });

    const service = React.useMemo(() => new ActivityTaskService(context), [context]);

    React.useEffect(() => {
        const loadData = async () => {
            setLoading(true);
            try {
                // Load lookups and choices
                const [acts, zones, profiles, cons, likes] = await Promise.all([
                    service.getLookupOptions('ActivityList'), // Example list name
                    service.getLookupOptions('WorkZoneList'),
                    service.getLookupOptions('BusinessProfileList'),
                    service.getChoiceOptions('Consequence'),
                    service.getChoiceOptions('Likelihood')
                ]);

                setLookups({ activities: acts, workZones: zones, businessProfiles: profiles });
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
                            onAdd={() => {}} // TODO: Implement add hazard
                            onDelete={() => {}} // TODO: Implement delete hazard
                            onEdit={() => {}} // TODO: Implement edit hazard
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
        </div>
    );
};


export default ActivityTaskForm;
