import * as React from 'react';
import { Stack, Text, Icon } from '@fluentui/react';

interface ITimelineProps {
    history: any[];
}

const Timeline: React.FC<ITimelineProps> = ({ history }) => {
    return (
        <Stack tokens={{ childrenGap: 20 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Clock" styles={{ root: { fontSize: 18, color: '#0078d4' } }} />
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 'bold' } }}>TIMELINE</Text>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
                <Icon iconName="Add" styles={{ root: { fontSize: 16, color: '#0078d4' } }} />
                <Text variant="medium" styles={{ root: { color: '#0078d4', cursor: 'pointer' } }}>New Timeline</Text>
            </Stack>

            <Stack tokens={{ childrenGap: 20 }} styles={{ root: { paddingLeft: 10, borderLeft: '2px solid #edebe9', marginLeft: 10 } }}>
                {history.length > 0 ? history.map((item, index) => (
                    <Stack key={index} tokens={{ childrenGap: 4 }}>
                        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                            <div style={{ 
                                width: 32, 
                                height: 32, 
                                borderRadius: '50%', 
                                backgroundColor: '#107c10', 
                                display: 'flex', 
                                alignItems: 'center', 
                                justifyContent: 'center',
                                color: 'white',
                                fontSize: 12
                            }}>
                                {item.userInitials || 'MH'}
                            </div>
                            <Stack>
                                <Stack horizontal tokens={{ childrenGap: 4 }}>
                                    <Text variant="smallPlus" styles={{ root: { fontWeight: 'bold' } }}>{item.userName || 'Matthew Hellmich'}</Text>
                                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>- {item.date || '20 Feb 2025, 03:28 PM'}</Text>
                                </Stack>
                                <Text variant="smallPlus">{item.description || 'New Risk Register Item Added'}</Text>
                            </Stack>
                        </Stack>
                    </Stack>
                )) : (
                    <Text variant="small" styles={{ root: { fontStyle: 'italic', color: '#a19f9d' } }}>No timeline entries yet.</Text>
                )}
            </Stack>
        </Stack>
    );
};

export default Timeline;
