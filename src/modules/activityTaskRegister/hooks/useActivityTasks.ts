import { useState, useCallback, useEffect, useMemo } from 'react';
import { ActivityTaskService } from '../services/ActivityTaskService';
import { IActivityTaskItem } from '../models/IActivityTask';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export const useActivityTasks = (context: WebPartContext) => {
    const [items, setItems] = useState<IActivityTaskItem[]>([]);
    const [loading, setLoading] = useState(true);
    const [totalCount, setTotalCount] = useState(0);
    const [currentPage, setCurrentPage] = useState(1);
    const [searchQuery, setSearchQuery] = useState('');
    const [sortConfig, setSortConfig] = useState<{ field: string; isAscending: boolean }>({
        field: 'Id',
        isAscending: true
    });

    const service = useMemo(() => new ActivityTaskService(context), [context]);

    const PAGE_SIZE = 100;

    const fetchItems = useCallback(async () => {
        setLoading(true);
        try {
            const [data, total] = await Promise.all([
                service.getActivityTasksPaged(
                    currentPage,
                    PAGE_SIZE,
                    searchQuery,
                    sortConfig.field,
                    sortConfig.isAscending
                ),
                service.getTotalCount(searchQuery)
            ]);
            setItems(data);
            setTotalCount(total);
        } catch (error) {
            console.error("[useActivityTasks] Failed to fetch items", error);
        } finally {
            setLoading(false);
        }
    }, [service, currentPage, searchQuery, sortConfig]);

    useEffect(() => {
        fetchItems();
    }, [fetchItems]);

    const deleteItems = async (selectedIds: number[]) => {
        if (confirm(`Are you sure you want to delete ${selectedIds.length} item(s)?`)) {
            setLoading(true);
            try {
                for (const id of selectedIds) {
                    await service.deleteActivityTask(id);
                }
                await fetchItems();
            } catch (error) {
                console.error("[useActivityTasks] Delete failed", error);
                alert("Failed to delete items.");
            } finally {
                setLoading(false);
            }
        }
    };

    return {
        items,
        loading,
        totalCount,
        currentPage,
        setCurrentPage,
        searchQuery,
        setSearchQuery,
        sortConfig,
        setSortConfig,
        deleteItems,
        fetchItems,
        PAGE_SIZE
    };
};
