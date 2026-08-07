import * as React from 'react';
import styles from './App.module.scss';

export interface IKpiTileProps {
    count: number;
    label: string;
    actionLabel: string;
    color: string;
    onClick: () => void;
}

const KpiTile: React.FC<IKpiTileProps> = ({ count, label, actionLabel, color, onClick }) => {
    return (
        <div className={styles.kpiTile} onClick={onClick} role="button" tabIndex={0}>
            <span className={styles.kpiNumber} style={{ color }}>{count}</span>
            <span className={styles.kpiLabel}>{label}</span>
            <span className={styles.kpiAction} style={{ color }}>► {actionLabel}</span>
        </div>
    );
};

export default KpiTile;
