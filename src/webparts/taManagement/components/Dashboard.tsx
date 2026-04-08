import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, IKpiData, AppView } from '../models/types';
import KpiTile from './KpiTile';
import StatusPill from './StatusPill';

export interface IDashboardProps {
    tas: ITaItem[];
    kpi: IKpiData;
    userName: string;
    onNavigate: (view: AppView, filterStatus?: string) => void;
    onSelectTa: (id: number) => void;
}

const Dashboard: React.FC<IDashboardProps> = ({ tas, kpi, userName, onNavigate, onSelectTa }) => {
    const recentTas = tas
        .filter(t => t.Status !== 'abgeschlossen')
        .sort((a, b) => new Date(b.Modified || '').getTime() - new Date(a.Modified || '').getTime())
        .slice(0, 5);

    return (
        <>
            {/* Header */}
            <div className={styles.header}>
                <span className={styles.headerTitle}>TA-Management</span>
                <span className={styles.headerRight}>
                    Willkommen, {userName}
                </span>
            </div>

            {/* Content */}
            <div className={styles.content}>
                {/* KPI Tiles */}
                <div className={styles.kpiRow}>
                    <KpiTile
                        count={kpi.overdue}
                        label="Überfällig"
                        actionLabel="Anzeigen"
                        color="#EF4444"
                        onClick={() => onNavigate(AppView.AlleTas, 'überfällig')}
                    />
                    <KpiTile
                        count={kpi.plan}
                        label="Termin planen"
                        actionLabel="Planen"
                        color="#3B82F6"
                        onClick={() => onNavigate(AppView.TerminPlanen)}
                    />
                    <KpiTile
                        count={kpi.onTrack}
                        label="Planmäßig"
                        actionLabel="Anzeigen"
                        color="#10B981"
                        onClick={() => onNavigate(AppView.AlleTas, 'läuft planmäßig')}
                    />
                    <KpiTile
                        count={kpi.review}
                        label="Prüfen"
                        actionLabel="Anzeigen"
                        color="#F59E0B"
                        onClick={() => onNavigate(AppView.AlleTas, 'prüfen')}
                    />
                </div>

                {/* Action Tiles */}
                <div className={styles.actionRow}>
                    <div
                        className={styles.actionTileBlue}
                        onClick={() => onNavigate(AppView.NeueTa)}
                        role="button"
                        tabIndex={0}
                    >
                        <span className={styles.actionTitle} style={{ color: '#3B82F6' }}>
                            Neue TA anlegen
                        </span>
                        <span className={styles.actionSub}>Technische Anfrage schnell erfassen</span>
                    </div>
                    <div
                        className={styles.actionTileGreen}
                        onClick={() => onNavigate(AppView.TerminPlanen)}
                        role="button"
                        tabIndex={0}
                    >
                        <span className={styles.actionTitle} style={{ color: '#10B981' }}>
                            Nächste Termine
                        </span>
                        <span className={styles.actionSub}>{kpi.plan} Termine offen</span>
                    </div>
                </div>

                {/* Recent Activity */}
                <h3 className={styles.sectionHeader}>Letzte Aktivitäten</h3>
                {recentTas.length === 0 ? (
                    <div className={styles.emptyState}>
                        <span className={styles.emptyIcon}></span>
                        <span>Keine offenen TAs vorhanden</span>
                    </div>
                ) : (
                    recentTas.map(ta => (
                        <div
                            key={ta.ID}
                            className={styles.listRow}
                            onClick={() => onSelectTa(ta.ID)}
                        >
                            <span className={styles.listRowBold}>{ta.Title}</span>
                            <span className={styles.listRowMuted}>{ta.field_8}</span>
                            <span className={styles.listRowMuted}>{ta.field_12}</span>
                            <span className={styles.listRowMuted} title={ta.Verantwortlicher?.Title}>{ta.Verantwortlicher?.Title || '–'}</span>
                            <StatusPill status={ta.Status || ''} />
                        </div>
                    ))
                )}
            </div>
        </>
    );
};

export default Dashboard;
