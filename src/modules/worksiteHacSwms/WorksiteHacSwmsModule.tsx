import * as React from 'react';
import { IColumn, Icon, Panel, PanelType } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import GenericGrid from '../../webparts/swf/components/Shared/GenericGrid';
import WorksiteHacSwmsForm from './components/WorksiteHacSwmsForm';
import { worksiteHacSwmsRecords } from './data/mockWorksiteHacSwmsData';
import { IWorksiteHacSwmsRecord } from './models/IWorksiteHacSwmsRecord';

export interface IWorksiteHacSwmsModuleProps {
    context: WebPartContext;
}

const formatDate = (dateValue: string): string => {
    const date = new Date(dateValue);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${day}/${month}/${date.getFullYear()}`;
};

const createDraftRecord = (): IWorksiteHacSwmsRecord => ({
    ...worksiteHacSwmsRecords[0],
    id: 0,
    number: 'New',
    date: new Date().toISOString(),
    project: '',
    scopeOfWorks: '',
    workAddresses: '',
    status: 'Draft',
    swmsUsed: [],
    timeline: []
});

const WorksiteHacSwmsModule: React.FC<IWorksiteHacSwmsModuleProps> = () => {
    const [selectedItem, setSelectedItem] = React.useState<IWorksiteHacSwmsRecord | undefined>(undefined);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);

    const columns: IColumn[] = [
        {
            key: 'number',
            name: 'ID',
            fieldName: 'number',
            minWidth: 80,
            maxWidth: 100,
            isResizable: true,
            onRender: (item?: IWorksiteHacSwmsRecord) => <a>{item?.number}</a>
        },
        {
            key: 'date',
            name: 'Date',
            fieldName: 'date',
            minWidth: 100,
            maxWidth: 120,
            isResizable: true,
            onRender: (item?: IWorksiteHacSwmsRecord) => <span>{item ? formatDate(item.date) : ''}</span>
        },
        {
            key: 'project',
            name: 'Project',
            fieldName: 'project',
            minWidth: 220,
            maxWidth: 320,
            isResizable: true
        },
        {
            key: 'scopeOfWorks',
            name: 'Scope Of Work',
            fieldName: 'scopeOfWorks',
            minWidth: 220,
            maxWidth: 320,
            isResizable: true
        },
        {
            key: 'submittedBy',
            name: 'Submitted By',
            fieldName: 'submittedBy',
            minWidth: 180,
            maxWidth: 220,
            isResizable: true,
            onRender: (item?: IWorksiteHacSwmsRecord) => {
                if (!item?.submittedBy) return <span>-</span>;
                return (
                    <span>
                        <Icon iconName="Contact" style={{ color: '#0078d4', marginRight: 6 }} />
                        {item.submittedBy}
                    </span>
                );
            }
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'status',
            minWidth: 140,
            maxWidth: 180,
            isResizable: true,
            onRender: (item?: IWorksiteHacSwmsRecord) => <span>{item?.status}</span>
        }
    ];

    const openForm = (item: IWorksiteHacSwmsRecord): void => {
        setSelectedItem(item);
        setIsPanelOpen(true);
    };

    const openNewForm = (): void => {
        setSelectedItem(createDraftRecord());
        setIsPanelOpen(true);
    };

    return (
        <React.Fragment>
            <GenericGrid
                items={worksiteHacSwmsRecords}
                columns={columns}
                loading={false}
                onNew={openNewForm}
                onEdit={openForm}
                onDelete={() => undefined}
                onRefresh={() => undefined}
                onExportExcel={() => undefined}
                onExportCSV={() => undefined}
                onExportPDF={() => undefined}
                onExportZip={() => undefined}
                currentPage={1}
                totalPages={1}
                totalCount={worksiteHacSwmsRecords.length}
                pageSize={50}
                clientSidePagination={true}
                onPageChange={() => undefined}
            />

            {selectedItem && (
                <Panel
                    isOpen={isPanelOpen}
                    onDismiss={() => setIsPanelOpen(false)}
                    type={PanelType.custom}
                    customWidth="calc(100vw - 96px)"
                    onRenderHeader={() => null}
                    isLightDismiss={false}
                    styles={{
                        content: { padding: 0 },
                        scrollableContent: { overflow: 'hidden' },
                        commands: { display: 'none' }
                    }}
                >
                    <WorksiteHacSwmsForm
                        item={selectedItem}
                        onClose={() => setIsPanelOpen(false)}
                    />
                </Panel>
            )}
        </React.Fragment>
    );
};

export default WorksiteHacSwmsModule;
