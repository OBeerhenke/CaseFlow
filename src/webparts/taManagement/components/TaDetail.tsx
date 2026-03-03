import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, STATUS_VALUES, INITIAL_DELAY_REASONS } from '../models/types';
import { SharePointService } from '../services/SharePointService';
import StatusPill from './StatusPill';
import Modal from './Modal';

export interface ITaDetailProps {
    ta: ITaItem;
    onBack: () => void;
    onSave: (id: number, fields: Partial<Record<string, string | number | undefined>>) => Promise<void>;
    onVerschieben: (id: number, neuerTermin: string, grund: string, alterTermin: string) => Promise<void>;
    onSetTermin?: (id: number, termin: string, verantwortlicherId?: number, delayReason?: string) => Promise<void>;
    onSearchUsers?: (query: string) => Promise<any[]>;
    onEnsureUser?: (login: string) => Promise<number>;
}

interface ITaDetailState {
    bemerkung: string;
    status: string;
    showModal: boolean;
    saving: boolean;
    termin: string;
    verantwortlicherSearchText: string;
    verantwortlicherLoginName: string | null;
    showVerantwortlicherSuggestions: boolean;
    verantwortlicherSuggestions: any[];
    savingTermin: boolean;
    showDelayModal: boolean;
    delayReason: string;
    delayThresholdDays: number;
}

export default class TaDetail extends React.Component<ITaDetailProps, ITaDetailState> {
    constructor(props: ITaDetailProps) {
        super(props);
        this.state = {
            bemerkung: props.ta.field_2 || '',
            status: props.ta.Status || '',
            showModal: false,
            saving: false,
            termin: '',
            verantwortlicherSearchText: '',
            verantwortlicherLoginName: null,
            showVerantwortlicherSuggestions: false,
            verantwortlicherSuggestions: [],
            savingTermin: false,
            showDelayModal: false,
            delayReason: '',
            delayThresholdDays: 2
        };
    }

    public async componentDidMount(): Promise<void> {
        try {
            const delayRaw = await SharePointService.instance.getConfigValue('DelayThresholdDays', '2');
            this.setState({ delayThresholdDays: parseInt(delayRaw, 10) || 2 });
        } catch (e) {
            console.error("Failed to load delay threshold config", e);
        }
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

    private handleVerantwortlicherChange = async (value: string): Promise<void> => {
        this.setState({
            verantwortlicherSearchText: value,
            verantwortlicherLoginName: null
        });

        if (value.length >= 3 && this.props.onSearchUsers) {
            const results = await this.props.onSearchUsers(value);
            this.setState({
                verantwortlicherSuggestions: results,
                showVerantwortlicherSuggestions: results.length > 0
            });
        } else {
            this.setState({
                verantwortlicherSuggestions: [],
                showVerantwortlicherSuggestions: false
            });
        }
    }

    private selectVerantwortlicher = (user: any): void => {
        this.setState({
            verantwortlicherSearchText: user.DisplayText,
            verantwortlicherLoginName: user.Key,
            showVerantwortlicherSuggestions: false
        });
    }

    private handlePlanen = async (): Promise<void> => {
        if (!this.state.termin || !this.props.onSetTermin) return;

        const planDate = new Date(this.state.termin); // e.g. "2026-03-05" from input type="date"
        let createdDate = new Date();
        if (this.props.ta.field_4) {
            const cd = new Date(this.props.ta.field_4);
            if (!isNaN(cd.getTime())) {
                createdDate = cd;
            } else {
                const pts = this.props.ta.field_4.split('.');
                if (pts.length === 3) createdDate = new Date(`${pts[2]}-${pts[1]}-${pts[0]}`);
            }
        }

        const today = new Date();

        // Compare midnight to midnight accurately
        const todayTime = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
        const createdTime = Date.UTC(createdDate.getFullYear(), createdDate.getMonth(), createdDate.getDate());
        const diffDays = (todayTime - createdTime) / (1000 * 3600 * 24);

        if (diffDays >= this.state.delayThresholdDays && !this.state.delayReason) {
            this.setState({ showDelayModal: true });
            return;
        }

        this.setState({ savingTermin: true });
        const parts = this.state.termin.split('-');
        const formatted = `${parts[2]}.${parts[1]}.${parts[0]}`;

        let finalVerantwortlicherId: number | undefined = undefined;
        if (this.state.verantwortlicherLoginName && this.props.onEnsureUser) {
            try {
                finalVerantwortlicherId = await this.props.onEnsureUser(this.state.verantwortlicherLoginName);
            } catch (e) {
                console.error("User not found or ensure failed", e);
            }
        }

        try {
            await this.props.onSetTermin(this.props.ta.ID, formatted, finalVerantwortlicherId, this.state.delayReason);
        } finally {
            this.setState({ savingTermin: false, verantwortlicherSearchText: '', verantwortlicherLoginName: null, termin: '', showDelayModal: false, delayReason: '' });
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

                        {!ta.field_6 ? (
                            <div style={{ marginTop: '16px', padding: '16px', background: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                                <h5 style={{ margin: '0 0 12px 0', color: '#334155' }}>Neuen Termin einplanen</h5>
                                <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'flex-end' }}>

                                    {/* Verantwortlicher Search */}
                                    <div style={{ flex: '1 1 200px', position: 'relative' }}>
                                        <label className={styles.formLabel} style={{ fontSize: '12px', marginBottom: '4px' }}>Verantwortlicher (optional)</label>
                                        <input
                                            className={styles.formInput}
                                            style={{ width: '100%' }}
                                            value={this.state.verantwortlicherSearchText}
                                            onChange={(e) => this.handleVerantwortlicherChange(e.target.value)}
                                            placeholder="Suchen (min. 3 Zeichen)"
                                        />
                                        {this.state.showVerantwortlicherSuggestions && (
                                            <div style={{ position: 'absolute', zIndex: 100, background: 'white', border: '1px solid #ccc', borderRadius: 4, maxHeight: 150, overflow: 'auto', width: '100%', top: 'calc(100% + 4px)', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}>
                                                {this.state.verantwortlicherSuggestions.map(u => (
                                                    <div
                                                        key={u.Key}
                                                        onClick={() => this.selectVerantwortlicher(u)}
                                                        style={{ padding: '6px 10px', cursor: 'pointer', fontSize: 13, color: 'black' }}
                                                        onMouseEnter={(e) => (e.currentTarget.style.background = '#f0f0f0')}
                                                        onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                                                    >
                                                        {u.DisplayText}
                                                    </div>
                                                ))}
                                            </div>
                                        )}
                                    </div>

                                    {/* Date Picker */}
                                    <div style={{ flex: '0 1 160px' }}>
                                        <label className={styles.formLabel} style={{ fontSize: '12px', marginBottom: '4px' }}>Datum</label>
                                        <input
                                            type="date"
                                            className={styles.formInput}
                                            style={{ width: '100%' }}
                                            value={this.state.termin}
                                            onChange={(e) => this.setState({ termin: e.target.value })}
                                        />
                                    </div>

                                    <button
                                        className={styles.btnSuccess}
                                        style={{ height: '36px', padding: '0 16px', whiteSpace: 'nowrap' }}
                                        disabled={!this.state.termin || this.state.savingTermin}
                                        onClick={this.handlePlanen}
                                    >
                                        {this.state.savingTermin ? 'Speichere...' : 'Termin einplanen'}
                                    </button>
                                </div>
                            </div>
                        ) : (
                            <button
                                className={styles.btnWarning}
                                onClick={() => this.setState({ showModal: true })}
                            >
                                Termin verschieben
                            </button>
                        )}
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

                    {/* Aktionen */}
                    {this.state.status !== 'abgeschlossen' && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>TA Abschließen</h4>
                            <p style={{ fontSize: 13, color: '#64748b', margin: '4px 0 12px 0' }}>
                                Der Status (z.B. prüfen, überfällig) wird von der App nun automatisch anhand des geplanten Termins berechnet.
                            </p>
                            <button
                                type="button"
                                className={styles.btnPrimary}
                                style={{ backgroundColor: '#10B981', borderColor: '#10B981' }}
                                onClick={() => {
                                    if (window.confirm("Bist du sicher, dass du diese TA abschließen möchtest? (Status wird auf 'abgeschlossen' gesetzt)")) {
                                        this.setState({ status: 'abgeschlossen' }, () => {
                                            void this.handleSave();
                                        });
                                    }
                                }}
                                disabled={this.state.saving}
                            >
                                TA erfolgreich abschließen
                            </button>
                        </div>
                    )}
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
                        {this.state.saving ? 'Speichere...' : 'Speichern'}
                    </button>
                </div>

                {/* Modal */}
                <Modal
                    visible={this.state.showModal}
                    onClose={() => this.setState({ showModal: false })}
                    onConfirm={this.handleVerschieben}
                />
                {/* Delay Modal */}
                {this.state.showDelayModal && (
                    <div className={styles.overlay} onClick={() => this.setState({ showDelayModal: false, delayReason: '' })}>
                        <div className={styles.modal} onClick={(e) => e.stopPropagation()}>
                            <h3 className={styles.modalTitle}>Verzögerungsgrund</h3>
                            <p style={{ fontSize: '13px', marginBottom: '16px', color: '#475569' }}>
                                Warum wurde die Fertigstellung der TA erst jetzt (mind. {this.state.delayThresholdDays} Tage nach der Erstellung) geplant?
                            </p>

                            <div className={styles.radioGroup}>
                                {INITIAL_DELAY_REASONS.map(grund => (
                                    <label key={grund} className={styles.radioItem}>
                                        <input
                                            type="radio"
                                            name="delayReason"
                                            value={grund}
                                            checked={this.state.delayReason === grund}
                                            onChange={() => this.setState({ delayReason: grund })}
                                        />
                                        {grund}
                                    </label>
                                ))}
                            </div>

                            <div className={styles.modalBtnRow} style={{ marginTop: '24px' }}>
                                <button className={styles.btnSecondary} onClick={() => this.setState({ showDelayModal: false, delayReason: '' })}>
                                    Abbrechen
                                </button>
                                <button
                                    className={styles.btnPrimary}
                                    disabled={!this.state.delayReason}
                                    onClick={() => this.handlePlanen()}
                                >
                                    Grund eintragen
                                </button>
                            </div>
                        </div>
                    </div>
                )}
            </>
        );
    }
}
