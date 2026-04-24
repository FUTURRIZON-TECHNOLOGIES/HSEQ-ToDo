import * as React from 'react';
import { IModuleInfo } from '../models/IModuleInfo';
import ToDoModule from '../../modules/todo/ToDoModule';
import TrainingInductionModule from '../../modules/trainingInduction';
import WorksiteHacSwmsModule, {
    WORKSITE_HAC_SWMS_MODULE_GROUP,
    WORKSITE_HAC_SWMS_MODULE_ICON,
    WORKSITE_HAC_SWMS_MODULE_ID,
    WORKSITE_HAC_SWMS_MODULE_LABEL
} from '../../modules/worksiteHacSwms';

export const ModuleRegistry: IModuleInfo[] = [
    {
        id: 'ToDo',
        label: 'To Do',
        iconName: 'TaskManager',
        group: 'HSEQ',
        component: ToDoModule
    },
    {
        id: 'TrainingInduction',
        label: 'Training & Inductions',
        iconName: 'Education',
        group: 'HSEQ',
        component: TrainingInductionModule
    },
    {
        id: WORKSITE_HAC_SWMS_MODULE_ID,
        label: WORKSITE_HAC_SWMS_MODULE_LABEL,
        iconName: WORKSITE_HAC_SWMS_MODULE_ICON,
        group: WORKSITE_HAC_SWMS_MODULE_GROUP,
        component: WorksiteHacSwmsModule
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
