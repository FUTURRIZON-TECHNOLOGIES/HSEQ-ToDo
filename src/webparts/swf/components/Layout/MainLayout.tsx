import * as React from 'react';
import styles from './MainLayout.module.scss';
import Sidebar from './Sidebar';
import { getModuleById } from '../../../../common/config/ModuleRegistry';

export interface IMainLayoutProps {
    activeModule: string;
    onModuleChange: (module: string) => void;
    children?: React.ReactNode;
}

const MainLayout: React.FC<IMainLayoutProps> = ({ children, activeModule, onModuleChange }) => {
    const [isMobileOpen, setIsMobileOpen] = React.useState(false);

    const handleItemClick = (module: string) => {
        onModuleChange(module);
        setIsMobileOpen(false);
    };

    const moduleInfo = getModuleById(activeModule);
    const label = moduleInfo ? moduleInfo.label : activeModule;

    return (
        <div className={styles.mainLayout}>
            {/* Mobile overlay */}
            {isMobileOpen && (
                <div
                    className={styles.overlay}
                    onClick={() => setIsMobileOpen(false)}
                />
            )}

            {/* ── Sidebar ── */}
            <div className={`${styles.sidebarWrapper} ${isMobileOpen ? styles.mobileOpen : ''}`}>
                <Sidebar
                    activeItem={activeModule}
                    onItemClick={handleItemClick}
                />
            </div>

            {/* ── Main Content Area ── */}
            <div className={styles.contentArea}>
                {/* Header */}
                <header className={styles.header}>
                    {/* Mobile hamburger — only visible on small screens */}
                    <button
                        className={styles.mobileMenuBtn}
                        onClick={() => setIsMobileOpen(!isMobileOpen)}
                        aria-label="Open navigation"
                    >
                        <span /><span /><span />
                    </button>

                    {/* Breadcrumb */}
                    <div className={styles.breadcrumb}>
                        <span className={styles.breadcrumbRoot}>ASP Assist Group</span>
                        <span className={styles.breadcrumbSep}>›</span>
                        <span className={styles.breadcrumbCurrent}>
                            {label}
                        </span>
                    </div>
                </header>

                {/* Page Content */}
                <div className={styles.mainContent}>
                    {children}
                </div>
            </div>
        </div>
    );
};

export default MainLayout;
