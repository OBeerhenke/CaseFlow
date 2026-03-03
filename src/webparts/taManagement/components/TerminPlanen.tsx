import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem } from '../models/types';
import StatusPill from './StatusPill';

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
                                className={styles.listRow}
                                onClick={() => this.props.onSelectTa(ta.ID)}
                                style={{ cursor: 'pointer' }}
                            >
                                <span className={styles.listRowBold}>{ta.Title}</span>
                                <span className={styles.listRowMuted}>{ta.field_8}</span>
                                <span className={styles.listRowMuted}>{ta.field_12}</span>
                                <span className={styles.listRowMuted} title={ta.Ersteller?.Title}>{ta.Ersteller?.Title || ''}</span>
                                <StatusPill status={ta.Status || ''} />
                            </div>
                        ))
                    )}
                </div>
            </>
        );
    }
}
