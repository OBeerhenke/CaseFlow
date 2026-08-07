import * as React from 'react';
import styles from './App.module.scss';

export interface IStatusPillProps {
    status: string;
}

const StatusPill: React.FC<IStatusPillProps> = ({ status }) => {
    const getClass = (): string => {
        switch (status) {
            case 'überfällig': return styles.statusOverdue;
            case 'Termin planen': return styles.statusPlan;
            case 'läuft planmäßig': return styles.statusOntrack;
            case 'prüfen': return styles.statusReview;
            case 'abgeschlossen': return styles.statusDone;
            default: return styles.statusPill;
        }
    };

    return <span className={getClass()}>{status}</span>;
};

export default StatusPill;
