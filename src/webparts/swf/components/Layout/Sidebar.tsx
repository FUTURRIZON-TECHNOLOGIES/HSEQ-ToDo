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
                {/* HSEQ Section */}
                <div className={styles.navGroup}>
                    {!isCollapsed && (
                        <div className={styles.groupHeader} onClick={() => setHseqExpanded(!hseqExpanded)}>
                            <Icon iconName="CheckList" className={styles.groupIcon} />
                            <span>HSEQ</span>
                            <Icon iconName={hseqExpanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                        </div>
                    )}
                    {(hseqExpanded || isCollapsed) && (
                        <div className={styles.subItems}>
                            <div
                                className={`${styles.navItem} ${activeItem === 'ToDo' ? styles.active : ''}`}
                                onClick={() => onItemClick('ToDo')}
                                title={isCollapsed ? 'To Do' : ''}
                            >
                                <Icon iconName="TaskManager" className={styles.itemIcon} />
                                {!isCollapsed && <span>To Do</span>}
                            </div>
                            <div
                                className={`${styles.navItem} ${activeItem === 'Compliance' ? styles.active : ''}`}
                                onClick={() => onItemClick('Compliance')}
                                title={isCollapsed ? 'Compliance' : ''}
                            >
                                <Icon iconName="ReadingMode" className={styles.itemIcon} />
                                {!isCollapsed && <span>Compliance Register</span>}
                            </div>
                            <div
                                className={`${styles.navItem} ${activeItem === 'TrainingInduction' ? styles.active : ''}`}
                                onClick={() => onItemClick('TrainingInduction')}
                                title={isCollapsed ? 'Training & Inductions' : ''}
                            >
                                <Icon iconName="Education" className={styles.itemIcon} />
                                {!isCollapsed && <span>Training & Inductions</span>}
                            </div>
                        </div>
                    )}
                </div>

                {/* Admin Section */}
                <div className={styles.navGroup}>
                    {!isCollapsed && (
                        <div className={styles.groupHeader} onClick={() => setAdminExpanded(!adminExpanded)}>
                            <Icon iconName="Admin" className={styles.groupIcon} />
                            <span>Admin</span>
                            <Icon iconName={adminExpanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                        </div>
                    )}
                    {(adminExpanded || isCollapsed) && (
                        <div className={styles.subItems}>
                            <div
                                className={`${styles.navItem} ${activeItem === 'Projects' ? styles.active : ''}`}
                                onClick={() => onItemClick('Projects')}
                                title={isCollapsed ? 'Projects' : ''}
                            >
                                <Icon iconName="ProjectCollection" className={styles.itemIcon} />
                                {!isCollapsed && <span>Projects</span>}
                            </div>
                        </div>
                    )}
                </div>
            </nav>
        </aside>
    );
};

export default Sidebar;
