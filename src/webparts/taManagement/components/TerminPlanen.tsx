import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem } from '../models/types';

export interface ITerminPlanenProps {
    tas: ITaItem[];
    onSetTermin: (id: number, termin: string, verantwortlicherId?: number) => Promise<void>;
    onBack: () => void;
    onSearchUsers?: (query: string) => Promise<any[]>;
    onEnsureUser?: (login: string) => Promise<number>;
}

interface ITerminPlanenState {
    termine: { [id: number]: string };
    verantwortlicherSearchText: { [id: number]: string };
    verantwortlicherLoginName: { [id: number]: string | null };
    showVerantwortlicherSuggestions: { [id: number]: boolean };
    verantwortlicherSuggestions: { [id: number]: any[] };
    savingId: number | undefined;
}

export default class TerminPlanen extends React.Component<ITerminPlanenProps, ITerminPlanenState> {
    constructor(props: ITerminPlanenProps) {
        super(props);
        this.state = {
            termine: {},
            verantwortlicherSearchText: {},
            verantwortlicherLoginName: {},
            showVerantwortlicherSuggestions: {},
            verantwortlicherSuggestions: {},
            savingId: undefined
        };
    }

    private getPlanTas(): ITaItem[] {
        return this.props.tas
            .filter(t => t.Status === 'Termin planen')
            .sort((a, b) => (a.field_13 || 99) - (b.field_13 || 99));
    }

    private handleVerantwortlicherChange = async (taId: number, value: string): Promise<void> => {
        this.setState({
            verantwortlicherSearchText: { ...this.state.verantwortlicherSearchText, [taId]: value },
            verantwortlicherLoginName: { ...this.state.verantwortlicherLoginName, [taId]: null }
        });

        if (value.length >= 3 && this.props.onSearchUsers) {
            const results = await this.props.onSearchUsers(value);
            this.setState({
                verantwortlicherSuggestions: { ...this.state.verantwortlicherSuggestions, [taId]: results },
                showVerantwortlicherSuggestions: { ...this.state.showVerantwortlicherSuggestions, [taId]: results.length > 0 }
            });
        } else {
            this.setState({
                verantwortlicherSuggestions: { ...this.state.verantwortlicherSuggestions, [taId]: [] },
                showVerantwortlicherSuggestions: { ...this.state.showVerantwortlicherSuggestions, [taId]: false }
            });
        }
    }

    private selectVerantwortlicher = (taId: number, user: any): void => {
        this.setState({
            verantwortlicherSearchText: { ...this.state.verantwortlicherSearchText, [taId]: user.DisplayText },
            verantwortlicherLoginName: { ...this.state.verantwortlicherLoginName, [taId]: user.Key },
            showVerantwortlicherSuggestions: { ...this.state.showVerantwortlicherSuggestions, [taId]: false }
        });
    }

    private handleSave = async (ta: ITaItem): Promise<void> => {
        const termin = this.state.termine[ta.ID];
        if (!termin) return;

        this.setState({ savingId: ta.ID });
        const parts = termin.split('-');
        const formatted = `${parts[2]}.${parts[1]}.${parts[0]}`;

        let finalVerantwortlicherId: number | undefined = undefined;
        const loginName = this.state.verantwortlicherLoginName[ta.ID];
        if (loginName && this.props.onEnsureUser) {
            try {
                finalVerantwortlicherId = await this.props.onEnsureUser(loginName);
            } catch (e) {
                console.error("User not found or ensure failed", e);
            }
        }

        try {
            await this.props.onSetTermin(ta.ID, formatted, finalVerantwortlicherId);
        } finally {
            this.setState({ savingId: undefined });
        }
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
                        planTas.map((ta, index) => (
                            <div key={ta.ID} className={styles.terminCard} style={{ position: 'relative', zIndex: planTas.length - index }}>
                                <div className={styles.terminCardHeader}>
                                    <span style={{ fontWeight: 600, color: 'white' }}>{ta.Title}</span>
                                    <span className={styles.terminCardMeta}>{ta.field_8}</span>
                                </div>
                                <div className={styles.terminCardMeta}>
                                    Material: {ta.field_12 || '–'} &nbsp;│&nbsp; Wunsch: {ta.field_5 || '–'}
                                </div>
                                <div className={styles.terminCardAction}>
                                    <div style={{ display: 'flex', gap: '10px', alignItems: 'center', marginBottom: '10px', flexWrap: 'nowrap' }}>
                                        <span style={{ color: 'white', fontSize: 13, flexShrink: 0, width: '130px' }}>Verantwortlicher:</span>
                                        <div style={{ position: 'relative', flex: 1 }}>
                                            <input
                                                className={styles.dateInput}
                                                style={{ width: '100%', textOverflow: 'ellipsis' }}
                                                value={this.state.verantwortlicherSearchText[ta.ID] || ''}
                                                onChange={(e) => this.handleVerantwortlicherChange(ta.ID, e.target.value)}
                                                placeholder="Suchen (min. 3 Zeichen)"
                                            />
                                            {this.state.showVerantwortlicherSuggestions[ta.ID] && (
                                                <div style={{ position: 'absolute', zIndex: 100, background: 'white', border: '1px solid #ccc', borderRadius: 4, maxHeight: 150, overflow: 'auto', width: '100%', top: 'calc(100% + 4px)', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}>
                                                    {this.state.verantwortlicherSuggestions[ta.ID]?.map(u => (
                                                        <div
                                                            key={u.Key}
                                                            onClick={() => this.selectVerantwortlicher(ta.ID, u)}
                                                            style={{ padding: '6px 10px', cursor: 'pointer', fontSize: 12, color: 'black' }}
                                                            onMouseEnter={(e) => (e.currentTarget.style.background = '#f0f0f0')}
                                                            onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                                                        >
                                                            {u.DisplayText}
                                                        </div>
                                                    ))}
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                    <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                                        <span style={{ color: 'white', fontSize: 13, flexShrink: 0, width: '130px' }}>Geplanter Termin:</span>
                                        <input
                                            type="date"
                                            className={styles.dateInput}
                                            style={{ flex: 1 }}
                                            value={this.state.termine[ta.ID] || ''}
                                            onChange={(e) => this.setState({
                                                termine: { ...this.state.termine, [ta.ID]: e.target.value }
                                            })}
                                        />
                                        <button
                                            className={styles.btnSuccess}
                                            disabled={!this.state.termine[ta.ID] || this.state.savingId === ta.ID}
                                            onClick={() => this.handleSave(ta)}
                                        >
                                            {this.state.savingId === ta.ID ? 'Speichern...' : '✓ Speichern'}
                                        </button>
                                    </div>
                                </div>
                            </div>
                        ))
                    )}
                </div>
            </>
        );
    }
}
