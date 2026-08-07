import * as React from 'react';
import styles from './App.module.scss';
import { AppView } from '../models/types';
import { Icon } from '@fluentui/react/lib/Icon';

export interface INavBarProps {
    activeView: AppView;
    onNavigate: (view: AppView) => void;
}

const NavBar: React.FC<INavBarProps> = ({ activeView, onNavigate }) => {
    const items = [
        { view: AppView.Dashboard, iconName: 'HomeSolid', label: 'Home' },
        { view: AppView.NewCase, iconName: 'Add', label: 'Neue TA' },
        { view: AppView.CaseList, iconName: 'BulletedList', label: 'Alle TAs' },
        { view: AppView.Analytics, iconName: 'BarChartVertical', label: 'Analytics' },
        { view: AppView.Schedule, iconName: 'Calendar', label: 'Planen' },
        { view: AppView.Settings, iconName: 'Settings', label: 'Einstellungen' },
    ];

    return (
        <nav className={styles.navBar}>
            {items.map(item => (
                <button
                    key={item.view}
                    className={`${styles.navItem} ${activeView === item.view ? styles.navActive : styles.navInactive}`}
                    onClick={() => onNavigate(item.view)}
                >
                    <span className={styles.navIcon}>
                        <Icon iconName={item.iconName} styles={{ root: { fontSize: '20px', color: '#1e293b' } }} />
                    </span>
                    <span className={styles.navLabel}>{item.label}</span>
                </button>
            ))}
        </nav>
    );
};

export default NavBar;
