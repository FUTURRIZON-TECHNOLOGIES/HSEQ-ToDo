import * as React from 'react';
import { 
    Stack, 
    Text, 
    PrimaryButton, 
    IconButton,
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    SelectionMode,
    Dropdown,
    IDropdownOption,
    TextField
} from '@fluentui/react';
import { IActivityTaskHazard } from '../models/IActivityTask';

interface IHazardManagerProps {
    hazards: IActivityTaskHazard[];
    onAdd: () => void;
    onDelete: (id: number) => void;
    onEdit: (hazard: IActivityTaskHazard) => void;
    revisedAssessment: {
        consequence?: string;
        likelihood?: string;
        ranking?: string;
    };
    onRevisedChange: (field: string, value: any) => void;
    choices: {
        consequences: string[];
        likelihoods: string[];
    };
}

const HazardManager: React.FC<IHazardManagerProps> = ({
    hazards,
    onAdd,
    onDelete,
    onEdit,
    revisedAssessment,
    onRevisedChange,
    choices
}) => {
    const mapChoices = (items: string[]): IDropdownOption[] =>
        items.map(i => ({ key: i, text: i }));

    const columns: IColumn[] = [
        {
            key: 'actions',
            name: '',
            fieldName: 'id',
            minWidth: 70,
            maxWidth: 70,
            onRender: (item: IActivityTaskHazard) => (
                <Stack horizontal>
                    <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => onEdit(item)} />
                    <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => onDelete(item.Id)} />
                </Stack>
            )
        },
        {
            key: 'hazard',
            name: 'Hazard',
            fieldName: 'Title',
            minWidth: 200,
            isResizable: true,
            onRender: (item: IActivityTaskHazard) => (
                <Text>{item.Title}</Text>
            )
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold' } }}>HAZARDS</Text>
                <PrimaryButton 
                    iconProps={{ iconName: 'Add' }} 
                    text="ADD NEW HAZARD" 
                    onClick={onAdd}
                />
            </Stack>

            <div style={{ border: '1px solid #edebe9', borderRadius: '2px' }}>
                <DetailsList
                    items={hazards}
                    columns={columns}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.none}
                />
            </div>

            <Stack tokens={{ childrenGap: 12 }}>
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold' } }}>Revised Risk Assessment</Text>
                <Stack horizontal tokens={{ childrenGap: 20 }}>
                    <Dropdown
                        label="Revised Consequence"
                        selectedKey={revisedAssessment.consequence}
                        options={mapChoices(choices.consequences)}
                        onChange={(_, opt) => onRevisedChange('RevisedConsequence', opt?.key)}
                        styles={{ root: { flex: 1 } }}
                    />
                    <Dropdown
                        label="Revised Likelihood"
                        selectedKey={revisedAssessment.likelihood}
                        options={mapChoices(choices.likelihoods)}
                        onChange={(_, opt) => onRevisedChange('RevisedLikelihood', opt?.key)}
                        styles={{ root: { flex: 1 } }}
                    />
                </Stack>
                <TextField
                    label="Revised Ranking"
                    readOnly
                    value={revisedAssessment.ranking || ''}
                />
            </Stack>
        </Stack>
    );
};

export default HazardManager;
