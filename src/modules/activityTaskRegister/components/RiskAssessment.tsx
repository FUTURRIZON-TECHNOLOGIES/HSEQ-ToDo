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
    };
    choices: {
        consequences: string[];
        likelihoods: string[];
    };
}

const RiskAssessment: React.FC<IRiskAssessmentProps> = ({ data, onChange, lookups, choices }) => {
    const mapToOptions = (items: ISPLookup[]): IDropdownOption[] =>
        items.map(i => ({ key: i.Id, text: i.Title }));

    const mapChoices = (items: string[]): IDropdownOption[] =>
        items.map(i => ({ key: i, text: i }));

    return (
        <Stack tokens={{ childrenGap: 14 }}>
            <Dropdown
                label="Business Profile"
                placeholder="-- Select Business Profile --"
                selectedKey={data.BusinessProfileId ?? null}
                options={mapToOptions(lookups.businessProfiles)}
                onChange={(_, opt) => onChange('BusinessProfileId', opt?.key)}
            />

            <Dropdown
                label="Activity"
                placeholder="-- Select Activity --"
                selectedKey={data.ActivityId ?? null}
                options={mapToOptions(lookups.activities)}
                onChange={(_, opt) => onChange('ActivityId', opt?.key)}
            />

            <Dropdown
                label="Work Zone"
                placeholder="-- Select Work Zone --"
                selectedKey={data.WorkZoneId ?? null}
                options={mapToOptions(lookups.workZones)}
                onChange={(_, opt) => onChange('WorkZoneId', opt?.key)}
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
                    placeholder="-- Select --"
                    selectedKey={data.Consequence ?? null}
                    options={mapChoices(choices.consequences)}
                    onChange={(_, opt) => onChange('Consequence', opt?.key)}
                    styles={{ root: { flex: 1 } }}
                />
                <Dropdown
                    label="Likelihood"
                    placeholder="-- Select --"
                    selectedKey={data.Likelihood ?? null}
                    options={mapChoices(choices.likelihoods)}
                    onChange={(_, opt) => onChange('Likelihood', opt?.key)}
                    styles={{ root: { flex: 1 } }}
                />
            </Stack>

            <TextField
                label="Risk Ranking"
                value={data.RiskRanking || ''}
                onChange={(_, val) => onChange('RiskRanking', val)}
            />

            <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="center">
                <Checkbox
                    label="High Risk Work"
                    checked={!!data.HighRiskWork}
                    onChange={(_, checked) => onChange('HighRiskWork', checked)}
                />
                <Checkbox
                    label="Active"
                    checked={data.Active !== false}
                    onChange={(_, checked) => onChange('Active', checked)}
                />
            </Stack>
        </Stack>
    );
};

export default RiskAssessment;
