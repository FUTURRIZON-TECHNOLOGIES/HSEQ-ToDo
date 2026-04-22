import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IColumn, Panel, PanelType } from '@fluentui/react';
import GenericGrid from '../../webparts/swf/components/Shared/GenericGrid';
import { useActivityTasks } from './hooks/useActivityTasks';
import ActivityTaskForm from './components/ActivityTaskForm';
import { IActivityTaskItem } from './models/IActivityTask';
import { ActivityTaskService } from './services/ActivityTaskService';

export interface IActivityTaskModuleProps {
    context: WebPartContext;
}

const ActivityTaskRegisterModule: React.FC<IActivityTaskModuleProps> = ({ context }) => {
    const {
        items,
        loading,
        totalCount,
        currentPage,
        setCurrentPage,
        setSearchQuery,
        sortConfig,
        setSortConfig,
        deleteItems,
        fetchItems,
        PAGE_SIZE
    } = useActivityTasks(context);

    const [selectedItem, setSelectedItem] = React.useState<IActivityTaskItem | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);

    const service = React.useMemo(() => new ActivityTaskService(context), [context]);

    const columns: IColumn[] = [
        {
            key: 'Id', name: 'ID', fieldName: 'Id',
            minWidth: 40, maxWidth: 60, isResizable: true
        },
        {
            key: 'Activity', name: 'Activity', fieldName: 'ActivityValue',
            minWidth: 120, maxWidth: 200, isResizable: true
        },
        {
            key: 'WorkZone', name: 'Work Zone', fieldName: 'WorkZoneValue',
            minWidth: 120, maxWidth: 200, isResizable: true
        },
        {
            key: 'Task', name: 'Task', fieldName: 'Task',
            minWidth: 200, maxWidth: 400, isResizable: true
        },
        {
            key: 'ResponsiblePersons', name: 'Responsible Persons', fieldName: 'ResponsiblePersonsValue',
            minWidth: 150, maxWidth: 250, isResizable: true
        },
        {
            key: 'RiskRanking', name: 'Risk Ranking', fieldName: 'RiskRanking',
            minWidth: 100, maxWidth: 150, isResizable: true
        },
        {
            key: 'Active', name: 'Active', fieldName: 'Active',
            minWidth: 60, maxWidth: 80, isResizable: true,
            onRender: (item: IActivityTaskItem) => item.Active ? 'Yes' : 'No'
        }
    ];

    const handleSave = async (payload: any, mode: 'stay' | 'close' | 'new') => {
        try {
            if (selectedItem?.Id) {
                await service.updateActivityTask(selectedItem.Id, payload);
            } else {
                await service.addActivityTask(payload);
            }

            await fetchItems();

            if (mode === 'close') {
                setIsPanelOpen(false);
            } else if (mode === 'new') {
                setSelectedItem(undefined);
            }
            // If stay, we might want to refresh the selected item but for now we'll keep it as is
        } catch (error) {
            console.error("Failed to save activity task", error);
            alert("Failed to save item.");
        }
    };

    return (
        <React.Fragment>
            <GenericGrid
                items={items}
                columns={columns}
                loading={loading}
                onNew={() => { setSelectedItem(undefined); setIsPanelOpen(true); }}
                onEdit={(item) => { setSelectedItem(item); setIsPanelOpen(true); }}
                onDelete={(selected) => deleteItems(selected.map(i => i.Id))}
                onRefresh={fetchItems}
                onSearch={setSearchQuery}
                currentPage={currentPage}
                totalPages={Math.ceil(totalCount / PAGE_SIZE)}
                totalCount={totalCount}
                pageSize={PAGE_SIZE}
                clientSidePagination={false}
                onPageChange={setCurrentPage}
                sortField={sortConfig.field}
                isAscending={sortConfig.isAscending}
                onSort={(field, isAsc) => setSortConfig({ field, isAscending: isAsc })}
            />

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                type={PanelType.custom}
                customWidth="1100px"
                onRenderHeader={() => null}
                isLightDismiss={false}
                styles={{
                    content: { padding: 0 },
                    scrollableContent: { overflow: 'hidden' },
                    commands: { display: 'none' }
                }}
            >
                <ActivityTaskForm
                    item={selectedItem}
                    context={context}
                    onClose={() => setIsPanelOpen(false)}
                    onSave={handleSave}
                    onRefresh={fetchItems}
                />
            </Panel>
        </React.Fragment>
    );
};

export default ActivityTaskRegisterModule;
