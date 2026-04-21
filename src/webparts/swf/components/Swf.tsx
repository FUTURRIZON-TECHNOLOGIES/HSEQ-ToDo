import * as React from 'react';
import MainLayout from './Layout/MainLayout';
import { ISwfProps } from './ISwfProps';
import { getModuleById } from '../../../common/config/ModuleRegistry';

const Swf: React.FC<ISwfProps> = ({ context }) => {
    const [activeModule, setActiveModule] = React.useState<string>('ToDo');

    const renderModule = () => {
        const moduleInfo = getModuleById(activeModule);
        
        if (moduleInfo && moduleInfo.component) {
            const ModuleComponent = moduleInfo.component;
            return <ModuleComponent context={context} />;
        }

        return (
            <div style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                height: '100%',
                flexDirection: 'column',
                gap: 12,
                color: '#666',
                fontSize: 16
            }}>
                <span style={{ fontSize: 40 }}>🚧</span>
                <span>This module is under development</span>
            </div>
        );
    };

    return (
        <MainLayout
            activeModule={activeModule}
            onModuleChange={setActiveModule}
        >
            {renderModule()}
        </MainLayout>
    );
};

export default Swf;
