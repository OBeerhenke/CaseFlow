import * as React from 'react';
import styles from './TaManagement.module.scss';
import { AppView } from '../models/types';

export interface INavBarProps {
    activeView: AppView;
    onNavigate: (view: AppView) => void;
}

const NavBar: React.FC<INavBarProps> = ({ activeView, onNavigate }) => {
    const items = [
        { view: AppView.Dashboard, icon: '', label: 'Home' },
        { view: AppView.NeueTa, icon: '', label: 'Neue TA' },
        { view: AppView.AlleTas, icon: '', label: 'Alle TAs' },
        { view: AppView.TerminPlanen, icon: '', label: 'Termine' },
    ];

    return (
        <nav className={styles.navBar}>
            {items.map(item => (
                <button
                    key={item.view}
                    className={`${styles.navItem} ${activeView === item.view ? styles.navActive : styles.navInactive}`}
                    onClick={() => onNavigate(item.view)}
                >
                    <span className={styles.navIcon}>{item.icon}</span>
                    <span className={styles.navLabel}>{item.label}</span>
                </button>
            ))}
        </nav>
    );
};

export default NavBar;
