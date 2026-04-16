import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem } from '../models/types';
import StatusPill from './StatusPill';

const prioLabel = (p?: number): { text: string; color: string } => {
    if (p === 3) return { text: '▲ Hoch', color: '#EF4444' };
    if (p === 2) return { text: '● Mittel', color: '#F59E0B' };
    if (p === 1) return { text: '▼ Niedrig', color: '#6B7280' };
    return { text: '–', color: '#CBD5E1' };
};

export interface ITerminPlanenProps {
    tas: ITaItem[];
    onSelectTa: (id: number) => void;
    onBack: () => void;
}

export default class TerminPlanen extends React.Component<ITerminPlanenProps> {
    private getPlanTas(): ITaItem[] {
        return this.props.tas
            .filter(t => t.Status === 'Termin planen')
            .sort((a, b) => (a.field_13 || 99) - (b.field_13 || 99));
    }

    public render(): React.ReactElement<ITerminPlanenProps> {
        const planTas = this.getPlanTas();

        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>Termin planen ({planTas.length} offen)</span>
                    </div>
                </div>

                <div className={styles.content}>
                    {planTas.length === 0 ? (
                        <div className={styles.emptyState}>
                            <span className={styles.emptyIcon}></span>
                            <span>Alle Termine sind geplant!</span>
                        </div>
                    ) : (
                        planTas.map((ta) => (
                            <div
                                key={ta.ID}
                                className={`${styles.listRow} ${styles.listRowTerminPlanen}`}
                                onClick={() => this.props.onSelectTa(ta.ID)}
                                style={{ cursor: 'pointer' }}
                            >
                                <span className={styles.listRowBold}>{ta.Title}</span>
                                <span className={styles.listRowMuted}>{ta.field_8 || '–'}</span>
                                <span className={styles.listRowMuted}>{ta.field_12 || '–'}</span>
                                <span className={styles.listRowMuted} title={ta.Verantwortlicher?.Title}>{ta.Verantwortlicher?.Title || '–'}</span>
                                <span style={{ fontSize: 11, fontWeight: 600, color: prioLabel(ta.field_13).color }}>{prioLabel(ta.field_13).text}</span>
                                <span className={styles.listRowWarn} title={ta.field_21}>{ta.field_21 ? `⚠ ${ta.field_21}` : ''}</span>
                                <span style={{ justifySelf: 'end' }}>
                                    <StatusPill status={ta.Status || ''} />
                                </span>
                            </div>
                        ))
                    )}
                </div>
            </>
        );
    }
}
