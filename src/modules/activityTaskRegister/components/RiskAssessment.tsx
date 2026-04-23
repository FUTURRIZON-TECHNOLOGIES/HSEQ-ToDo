import * as React from 'react';
import { 
    Dropdown, 
    TextField, 
    Checkbox, 
    IDropdownOption,
    Stack
} from '@fluentui/react';
import { ISPLookup } from '../models/IActivityTask';

interface IRiskAssessmentProps {
    data: any;
    onChange: (field: string, value: any) => void;
    lookups: {
        activities: ISPLookup[];
        workZones: ISPLookup[];
        businessProfiles: ISPLookup[];
        hazards: ISPLookup[];
    };
    choices: {
        consequences: string[];
        likelihoods: string[];
    };
}

const RiskAssessment: React.FC<IRiskAssessmentProps> = ({ 
    data, 
    onChange, 
    lookups, 
    choices 
}) => {
    const mapToOptions = (items: ISPLookup[]): IDropdownOption[] => 
        items.map(i => ({ key: i.Id, text: i.Title }));

    const mapChoices = (items: string[]): IDropdownOption[] =>
        items.map(i => ({ key: i, text: i }));

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Dropdown
                label="Business Profile"
                selectedKey={data.BusinessProfileId}
                options={mapToOptions(lookups.businessProfiles)}
                onChange={(_, opt) => onChange('BusinessProfileId', opt?.key)}
            />
            
            <Dropdown
                label="Activity"
                selectedKey={data.ActivityId}
                options={mapToOptions(lookups.activities)}
                onChange={(_, opt) => onChange('ActivityId', opt?.key)}
            />

            <Dropdown
                label="Work Zone"
                selectedKey={data.WorkZoneId}
                options={mapToOptions(lookups.workZones)}
                onChange={(_, opt) => onChange('WorkZoneId', opt?.key)}
            />

            <Dropdown
                label="Hazard"
                selectedKey={data.HazardId}
                options={mapToOptions(lookups.hazards)}
                onChange={(_, opt) => onChange('HazardId', opt?.key)}
            />

            <TextField
                label="Task"
                multiline
                rows={3}
                value={data.Task || ''}
                onChange={(_, val) => onChange('Task', val)}
            />

            <Stack horizontal tokens={{ childrenGap: 20 }}>
                <Dropdown
                    label="Consequence"
                    selectedKey={data.Consequence}
                    options={mapChoices(choices.consequences)}
                    onChange={(_, opt) => onChange('Consequence', opt?.key)}
                    styles={{ root: { flex: 1 } }}
                />
                <Dropdown
                    label="Likelihood"
                    selectedKey={data.Likelihood}
                    options={mapChoices(choices.likelihoods)}
                    onChange={(_, opt) => onChange('Likelihood', opt?.key)}
                    styles={{ root: { flex: 1 } }}
                />
            </Stack>

            <TextField
                label="Risk Ranking"
                readOnly
                value={data.RiskRanking || ''}
            />

            <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
                <Checkbox
                    label="High Risk Work"
                    checked={data.HighRiskWork}
                    onChange={(_, checked) => onChange('HighRiskWork', checked)}
                />
                <Checkbox
                    label="Active"
                    checked={data.Active}
                    onChange={(_, checked) => onChange('Active', checked)}
                />
            </Stack>
        </Stack>
    );
};

export default RiskAssessment;
