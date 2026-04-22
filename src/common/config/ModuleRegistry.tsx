import * as React from 'react';
import { IModuleInfo } from '../models/IModuleInfo';
import ToDoModule from '../../modules/todo/ToDoModule';

export const ModuleRegistry: IModuleInfo[] = [
    {
        id: 'ToDo',
        label: 'To Do',
        iconName: 'TaskManager',
        group: 'HSEQ',
        component: ToDoModule
    },
    {
        id: 'ActivityTaskRegister',
        label: 'Activity & Task Register',
        iconName: 'TaskGroup',
        group: 'HSEQ',
        component: React.lazy(() => import('../../modules/activityTaskRegister/ActivityTaskRegisterModule'))
    },
    {
        id: 'Compliance',
        label: 'Compliance Register',
        iconName: 'ReadingMode',
        group: 'HSEQ',
        component: () => (
            <div style={{ padding: 20 }}>
                <h2>Compliance Register</h2>
                <p>This module is currently under development.</p>
            </div>
        )
    },
    {
        id: 'Projects',
        label: 'Projects',
        iconName: 'ProjectCollection',
        group: 'Admin',
        component: () => (
            <div style={{ padding: 20 }}>
                <h2>Projects</h2>
                <p>This module is currently under development.</p>
            </div>
        )
    }
];

export const getModuleById = (id: string): IModuleInfo | undefined => {
    return ModuleRegistry.find(m => m.id === id);
};
