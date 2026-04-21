import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IModuleProps {
    context: WebPartContext;
}

export interface IModuleInfo {
    id: string;
    label: string;
    iconName: string;
    group: 'HSEQ' | 'Admin';
    component: React.FC<IModuleProps>;
}
