import * as React from 'react';
import {
    DetailsListLayoutMode,
    Selection,
    IColumn,
    SelectionMode,
    CommandBar,
    ICommandBarItemProps,
    ShimmeredDetailsList,
    ConstrainMode,
    Icon,
    SearchBox
} from '@fluentui/react';
import styles from './GenericGrid.module.scss';

export interface IGenericGridProps {
    items: any[];
    columns: IColumn[];
    onNew?: () => void;
    onDelete?: (selectedItems: any[]) => void;
    onExportExcel?: (items: any[], isSelection?: boolean) => void;
    onExportPDF?: (items: any[], isSelection?: boolean) => void;
    onExportCSV?: (items: any[], isSelection?: boolean) => void;
    onExportZip?: (items: any[], isSelection?: boolean) => void;
    onRefresh?: () => void;
    onEdit?: (item: any) => void;
    onSearch?: (term: string) => void;
    loading?: boolean;
    title?: string;
    currentPage?: number;
    totalPages?: number;
    totalCount?: number;
    pageSize?: number;
    onPageChange?: (page: number) => void;
    clientSidePagination?: boolean;

    // Sorting
    sortField?: string;
    isAscending?: boolean;
    onSort?: (field: string, isAsc: boolean) => void;
}

const GenericGrid: React.FC<IGenericGridProps> = (props) => {
    const [filteredItems, setFilteredItems] = React.useState<any[]>(props.items);
    const [selectionCount, setSelectionCount] = React.useState(0);
    const [searchQuery, setSearchQuery] = React.useState('');
    const listContainerRef = React.useRef<HTMLDivElement>(null);

    const [selection] = React.useState<Selection>(
        new Selection({
            onSelectionChanged: () => {
                setSelectionCount(selection.getSelectedCount());
            }
        })
    );

    const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        if (!props.onSort) return;
        const isAscending = props.sortField === column.key ? !props.isAscending : true;
        props.onSort(column.key, isAscending);
    };

    // Debounce search
    const onSearchRef = React.useRef(props.onSearch);
    React.useEffect(() => {
        onSearchRef.current = props.onSearch;
    }, [props.onSearch]);

    React.useEffect(() => {
        const handler = setTimeout(() => {
            onSearchRef.current?.(searchQuery);
        }, 500);

        return () => clearTimeout(handler);
    }, [searchQuery]);

    React.useEffect(() => {
        console.log(`[GenericGrid] Received ${props.items?.length || 0} items for grid:`, props.title);
        setFilteredItems(props.items);
        // Clear stale selection and scroll to top on every data change
        selection.setAllSelected(false);
        if (listContainerRef.current) {
            listContainerRef.current.scrollTop = 0;
        }
    }, [props.items]);

    // Clear selection on page change to prevent stale selection across pages
    React.useEffect(() => {
        selection.setAllSelected(false);
        if (listContainerRef.current) {
            listContainerRef.current.scrollTop = 0;
        }
    }, [props.currentPage]);

    const pagedItems = React.useMemo(() => {
        if (props.clientSidePagination && props.pageSize) {
            const start = ((props.currentPage ?? 1) - 1) * props.pageSize;
            const end = (props.currentPage ?? 1) * props.pageSize;
            return filteredItems.slice(start, end);
        }
        return filteredItems;
    }, [filteredItems, props.currentPage, props.pageSize, props.clientSidePagination]);

    // Update selection object whenever the displayed items change
    // This is CRITICAL for Selection to map indices correctly to items
    React.useEffect(() => {
        selection.setItems(pagedItems, true);
    }, [pagedItems]);

    const getExportTarget = (): any[] =>
        (selectionCount > 0 ? selection.getSelection() : filteredItems) || [];

    const commandItems: ICommandBarItemProps[] = [
        {
            key: 'new',
            text: 'New',
            iconProps: { iconName: 'Add' },
            className: styles.cmdNew,
            onClick: props.onNew
        },
        {
            key: 'edit',
            text: 'Edit',
            iconProps: { iconName: 'Edit' },
            disabled: selectionCount !== 1,
            onClick: (): void => {
                if (selectionCount === 1) props.onEdit?.(selection.getSelection()[0]);
            }
        },
        {
            key: 'delete',
            text: 'Delete',
            iconProps: { iconName: 'Delete' },
            disabled: selectionCount === 0,
            onClick: (): void => { if (props.onDelete) props.onDelete(selection.getSelection()); }
        },
        {
            key: 'export',
            text: 'Export',
            iconProps: { iconName: 'Download' },
            subMenuProps: {
                items: [
                    {
                        key: 'excel',
                        text: 'Excel (.xlsx)',
                        iconProps: { iconName: 'ExcelDocument' },
                        onClick: () => props.onExportExcel?.(getExportTarget(), selectionCount > 0)
                    },
                    {
                        key: 'csv',
                        text: 'CSV (.csv)',
                        iconProps: { iconName: 'TextDocument' },
                        onClick: () => props.onExportCSV?.(getExportTarget(), selectionCount > 0)
                    },
                    {
                        key: 'pdf',
                        text: 'PDF (.pdf)',
                        iconProps: { iconName: 'PDF' },
                        onClick: () => props.onExportPDF?.(getExportTarget(), selectionCount > 0)
                    },
                    { key: 'divider', itemType: 1 /* Divider */ },
                    {
                        key: 'zip',
                        text: 'ZIP — All Formats',
                        iconProps: { iconName: 'ZipFolder' },
                        onClick: () => props.onExportZip?.(getExportTarget(), selectionCount > 0)
                    }
                ]
            }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: props.onRefresh
        }
    ];

    const farItems: ICommandBarItemProps[] = [
        {
            key: 'search',
            onRender: (): JSX.Element => (
                <div className={styles.searchContainer}>
                    <SearchBox
                        placeholder="Search across all fields..."
                        className={styles.searchBox}
                        value={searchQuery}
                        onChange={(_, newValue) => setSearchQuery(newValue || '')}
                        onClear={() => setSearchQuery('')}
                        underlined={false}
                    />
                </div>
            )
        }
    ];

    // Enhance columns with sorting capability and indicators
    const styledColumns: IColumn[] = props.columns.map(col => ({
        ...col,
        isResizable: true,
        isSorted: props.sortField === col.key,
        isSortedDescending: props.sortField === col.key ? !props.isAscending : false,
        onColumnClick: onColumnClick
    }));
    
    // Clear selection when displayed pagedItems change (pagination or sorting)
    React.useEffect(() => {
        selection.setAllSelected(false);
        if (listContainerRef.current) listContainerRef.current.scrollTop = 0;
    }, [pagedItems, props.currentPage, props.sortField, props.isAscending]);

    const selectionLabel =
        selectionCount > 0
            ? `${selectionCount} of ${filteredItems.length} selected`
            : `${filteredItems.length} item${filteredItems.length !== 1 ? 's' : ''}`;

    return (
        <div className={styles.genericGrid}>
            {/* ── Toolbar ── */}
            <div className={styles.gridHeader}>
                <CommandBar
                    items={commandItems}
                    farItems={farItems}
                    className={styles.commandBar}
                    styles={{ root: { padding: '0 8px' } }}
                />
            </div>

            {/* ── Selection Summary ── */}
            {selectionCount > 0 && (
                <div className={styles.selectionBar}>
                    <div className={styles.selectionInfo}>
                        <Icon iconName="MultiSelect" className={styles.selectionIcon} />
                        <span className={styles.selectionBadge}>{selectionCount}</span>
                        {' '}item{selectionCount !== 1 ? 's' : ''} selected
                    </div>
                    <button
                        className={styles.clearSelection}
                        onClick={() => { selection.setAllSelected(false); }}
                    >
                        Clear Selection
                    </button>
                </div>
            )}

            {/* ── Data Grid ── */}
            <div className={styles.listContainer} ref={listContainerRef}>
                <ShimmeredDetailsList
                    items={pagedItems}
                    columns={styledColumns}
                    selection={selection}
                    selectionMode={SelectionMode.multiple}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    constrainMode={ConstrainMode.unconstrained}
                    enableShimmer={props.loading}
                    onItemInvoked={props.onEdit}
                    selectionPreservedOnEmptyClick={false}
                    onShouldVirtualize={() => false}
                    className={styles.detailsList}
                />
            </div>

            {/* ── Status Bar ── */}
            <div className={styles.statusBar}>
                <span>{selectionLabel}</span>
                {selectionCount > 0 && (
                    <span className={styles.exportHint}>
                        ↑ Export will use selected records
                    </span>
                )}
            </div>

            {/* ── Pagination Bar ── */}
            {props.onPageChange && (
                <div className={styles.paginationBar}>
                    <div className={styles.paginationLeft}>
                        <button
                            className={styles.pageBtn}
                            disabled={props.loading || (props.currentPage ?? 1) <= 1}
                            onClick={() => props.onPageChange!((props.currentPage ?? 1) - 1)}
                            title="Previous page"
                        >
                            ◀ Previous
                        </button>
                    </div>

                    <div className={styles.pageInfo}>
                        {props.loading ? (
                            <span className={styles.pageLoading}>Loading…</span>
                        ) : (
                            <span className={styles.pageTotal}>
                                Showing <strong>{((props.currentPage ?? 1) - 1) * 100 + 1}–{Math.min((props.currentPage ?? 1) * 100, props.totalCount || 0)}</strong> items
                            </span>
                        )}
                    </div>

                    <div className={styles.paginationRight}>
                        <button
                            className={styles.pageBtn}
                            disabled={props.loading || (props.currentPage ?? 1) >= (props.totalPages ?? 1)}
                            onClick={() => props.onPageChange!((props.currentPage ?? 1) + 1)}
                            title="Next page"
                        >
                            Next ▶
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
};

export default GenericGrid;
