import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, STATUS_VALUES } from '../models/types';
import StatusPill from './StatusPill';
import Modal from './Modal';

export interface ITaDetailProps {
    ta: ITaItem;
    onBack: () => void;
    onSave: (id: number, fields: Partial<Record<string, string | number | undefined>>) => Promise<void>;
    onVerschieben: (id: number, neuerTermin: string, grund: string, alterTermin: string) => Promise<void>;
}

interface ITaDetailState {
    bemerkung: string;
    status: string;
    showModal: boolean;
    saving: boolean;
}

export default class TaDetail extends React.Component<ITaDetailProps, ITaDetailState> {
    constructor(props: ITaDetailProps) {
        super(props);
        this.state = {
            bemerkung: props.ta.field_2 || '',
            status: props.ta.Status || '',
            showModal: false,
            saving: false
        };
    }

    private handleSave = async (): Promise<void> => {
        this.setState({ saving: true });
        try {
            await this.props.onSave(this.props.ta.ID, {
                field_2: this.state.bemerkung,
                Status: this.state.status
            });
        } finally {
            this.setState({ saving: false });
        }
    }

    private handleVerschieben = async (neuerTermin: string, grund: string): Promise<void> => {
        await this.props.onVerschieben(
            this.props.ta.ID,
            neuerTermin,
            grund,
            this.props.ta.field_6 || ''
        );
        this.setState({ showModal: false });
    }

    private formatCurrency(val: number | undefined): string {
        if (val === undefined || val === null) return '–';
        return val.toLocaleString('de-DE') + ' €';
    }

    public render(): React.ReactElement<ITaDetailProps> {
        const { ta } = this.props;

        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>{ta.Title}</span>
                    </div>
                </div>

                <div className={styles.content}>
                    {/* Status & Overview */}
                    <div className={styles.formGroup}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                            <StatusPill status={ta.Status || ''} />
                            <span style={{ color: 'rgb(100, 116, 139)', fontSize: 12 }}>
                                Ersteller: <strong>{ta.Ersteller?.Title || '–'}</strong>
                            </span>
                            <span style={{ color: 'rgb(100, 116, 139)', fontSize: 12 }}>
                                Verantwortlich: <strong>{ta.Verantwortlicher?.Title || '–'}</strong>
                            </span>
                            <span style={{ color: 'rgb(100, 116, 139)', fontSize: 12 }}>
                                Erstellt: {ta.field_4 || '–'}
                            </span>
                        </div>
                    </div>

                    {/* Termine */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Termine</h4>
                        <div className={styles.detailGrid}>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Wunschtermin</span>
                                <span className={styles.detailValue}>{ta.field_5 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Geplanter Termin</span>
                                <span className={styles.detailValue}>{ta.field_6 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Ursprünglicher Termin</span>
                                <span className={styles.detailValue}>{ta.field_22 || '–'}</span>
                            </div>
                        </div>
                        <button
                            className={styles.btnWarning}
                            onClick={() => this.setState({ showModal: true })}
                        >
                            📅 Termin verschieben
                        </button>
                    </div>

                    {/* Kundendaten */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Kundendaten</h4>
                        <div className={styles.detailGrid}>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Kunde</span>
                                <span className={styles.detailValue}>{ta.field_8 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Endkunde</span>
                                <span className={styles.detailValue}>{ta.field_9 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Projekt-Nr.</span>
                                <span className={styles.detailValue}>{ta.field_10 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Anwendung</span>
                                <span className={styles.detailValue}>{ta.field_11 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Material</span>
                                <span className={styles.detailValue}>{ta.field_12 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Kategorie</span>
                                <span className={styles.detailValue}>{ta.field_16 || '–'}</span>
                            </div>
                        </div>
                    </div>

                    {/* Kosten */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Kosten</h4>
                        <div className={styles.detailGrid}>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Budget bei Start</span>
                                <span className={styles.detailValue}>{ta.field_17 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Geplante Kosten</span>
                                <span className={styles.detailValue}>{this.formatCurrency(ta.field_19)}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>IST Kosten</span>
                                <span className={styles.detailValue}>{this.formatCurrency(ta.field_18)}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Differenz</span>
                                <span className={styles.detailValue} style={{
                                    color: ta.field_20 !== undefined && ta.field_20 >= 0 ? '#10B981' : '#EF4444'
                                }}>
                                    {this.formatCurrency(ta.field_20)}
                                </span>
                            </div>
                        </div>
                    </div>

                    {/* Bemerkungen */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Bemerkungen</h4>
                        <textarea
                            className={styles.formTextarea}
                            value={this.state.bemerkung}
                            onChange={(e) => this.setState({ bemerkung: e.target.value })}
                        />
                    </div>

                    {/* Verschiebungsgrund */}
                    {ta.field_21 && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>Letzter Verschiebungsgrund</h4>
                            <p style={{ fontSize: 13, color: '#F59E0B', margin: 0 }}>{ta.field_21}</p>
                        </div>
                    )}

                    {/* Status ändern */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Status ändern</h4>
                        <select
                            className={styles.formSelect}
                            value={this.state.status}
                            onChange={(e) => this.setState({ status: e.target.value })}
                        >
                            {STATUS_VALUES.map(s => (
                                <option key={s} value={s}>{s}</option>
                            ))}
                        </select>
                    </div>
                </div>

                {/* Footer */}
                <div className={styles.btnRow}>
                    <button className={styles.btnSecondary} onClick={this.props.onBack}>
                        ← Zurück
                    </button>
                    <button
                        className={styles.btnPrimary}
                        onClick={this.handleSave}
                        disabled={this.state.saving}
                    >
                        {this.state.saving ? '⏳ Speichere...' : '💾 Speichern'}
                    </button>
                </div>

                {/* Modal */}
                <Modal
                    visible={this.state.showModal}
                    onClose={() => this.setState({ showModal: false })}
                    onConfirm={this.handleVerschieben}
                />
            </>
        );
    }
}
