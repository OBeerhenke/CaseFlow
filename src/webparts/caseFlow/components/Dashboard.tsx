import * as React from 'react';
import styles from './App.module.scss';
import { ICaseItem, IKpiData, AppView } from '../models/types';
import { StatusConfigService } from '../services/StatusConfigService';
import { LabelService } from '../services/LabelService';
import { ThemeService } from '../services/ThemeService';
import KpiTile from './KpiTile';

export interface IDashboardProps {
    cases: ICaseItem[];
    kpi: IKpiData;
    config: Record<string, string>;
    userName: string;
    onNavigate: (view: AppView, filterStatus?: string) => void;
    onSelectCase: (id: number) => void;
}

const Dashboard: React.FC<IDashboardProps> = ({ cases, kpi, config, userName, onNavigate }) => {

    return (
        <>
            {/* Header */}
            <div className={styles.header}>
                <div className={styles.headerLeft}>
                    {ThemeService.getLogoUrl(config) && (
                        <img src={ThemeService.getLogoUrl(config)} alt="Logo" className={styles.headerLogo} />
                    )}
                    <span className={styles.headerTitle}>CaseFlow</span>
                </div>
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
                        label={StatusConfigService.getLabel('überfällig')}
                        actionLabel="Anzeigen"
                        color={StatusConfigService.getColor('überfällig')}
                        onClick={() => onNavigate(AppView.CaseList, 'überfällig')}
                    />
                    <KpiTile
                        count={kpi.plan}
                        label={StatusConfigService.getLabel('Termin planen')}
                        actionLabel="Planen"
                        color={StatusConfigService.getColor('Termin planen')}
                        onClick={() => onNavigate(AppView.Schedule)}
                    />
                    <KpiTile
                        count={kpi.onTrack}
                        label="Planmäßig"
                        actionLabel="Anzeigen"
                        color={StatusConfigService.getColor('läuft planmäßig')}
                        onClick={() => onNavigate(AppView.CaseList, 'läuft planmäßig')}
                    />
                    <KpiTile
                        count={kpi.review}
                        label={StatusConfigService.getLabel('prüfen')}
                        actionLabel="Anzeigen"
                        color={StatusConfigService.getColor('prüfen')}
                        onClick={() => onNavigate(AppView.CaseList, 'prüfen')}
                    />
                </div>

                {/* Action Tiles */}
                <div className={styles.actionRow}>
                    <div
                        className={styles.actionTileBlue}
                        onClick={() => onNavigate(AppView.NewCase)}
                        role="button"
                        tabIndex={0}
                    >
                        <span className={styles.actionTitle} style={{ color: StatusConfigService.getColor('Termin planen') }}>
                            {LabelService.getCreateActionLabel(config)}
                        </span>
                        <span className={styles.actionSub}>{LabelService.getEntityLabelSingular(config)} schnell erfassen</span>
                    </div>
                    <div
                        className={styles.actionTileGreen}
                        onClick={() => onNavigate(AppView.Schedule)}
                        role="button"
                        tabIndex={0}
                    >
                        <span className={styles.actionTitle} style={{ color: '#10B981' }}>
                            Nächste Termine
                        </span>
                        <span className={styles.actionSub}>{kpi.plan} Termine offen</span>
                    </div>
                </div>
            </div>
        </>
    );
};

export default Dashboard;
