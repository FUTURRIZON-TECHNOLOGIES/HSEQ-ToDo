import * as React from 'react';
import styles from './Sidebar.module.scss';
import { Icon } from '@fluentui/react';
import { ModuleRegistry } from '../../../../common/config/ModuleRegistry';

export interface ISidebarProps {
    activeItem: string;
    onItemClick: (item: string) => void;
    isCollapsed?: boolean;
}

const Sidebar: React.FC<ISidebarProps> = ({ activeItem, onItemClick, isCollapsed }) => {
    const [hseqExpanded, setHseqExpanded] = React.useState(true);
    const [adminExpanded, setAdminExpanded] = React.useState(true);

    const renderNavGroup = (title: string, groupKey: 'HSEQ' | 'Admin', expanded: boolean, setExpanded: (val: boolean) => void, icon: string) => {
        const groupModules = ModuleRegistry.filter(m => m.group === groupKey);

        return (
            <div className={styles.navGroup}>
                {!isCollapsed && (
                    <div className={styles.groupHeader} onClick={() => setExpanded(!expanded)}>
                        <Icon iconName={icon} className={styles.groupIcon} />
                        <span>{title}</span>
                        <Icon iconName={expanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                    </div>
                )}
                {(expanded || isCollapsed) && (
                    <div className={styles.subItems}>
                        {groupModules.map(module => (
                            <div 
                                key={module.id}
                                className={`${styles.navItem} ${activeItem === module.id ? styles.active : ''}`}
                                onClick={() => onItemClick(module.id)}
                                title={isCollapsed ? module.label : ''}
                            >
                                <Icon iconName={module.iconName} className={styles.itemIcon} />
                                {!isCollapsed && <span>{module.label}</span>}
                            </div>
                        ))}
                    </div>
                )}
            </div>
        );
    };

    return (
        <aside className={`${styles.sidebar} ${isCollapsed ? styles.collapsed : ''}`}>
            <div className={styles.logoSection}>
                <div className={styles.logo}>
                    {isCollapsed ? 'A' : 'ASP Assist Group'}
                </div>
            </div>

            <nav className={styles.nav}>
                {renderNavGroup('HSEQ', 'HSEQ', hseqExpanded, setHseqExpanded, 'CheckList')}
                {renderNavGroup('Admin', 'Admin', adminExpanded, setAdminExpanded, 'Admin')}
            </nav>
        </aside>
    );
};

export default Sidebar;
