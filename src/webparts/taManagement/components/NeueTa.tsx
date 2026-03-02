import * as React from 'react';
import styles from './TaManagement.module.scss';
import { INewTaForm, IProjektItem } from '../models/types';

export interface INeueTaProps {
    projekte: IProjektItem[];
    users: any[];
    onSubmit: (form: INewTaForm) => Promise<void>;
    onBack: () => void;
    onSearchUsers?: (query: string) => Promise<any[]>;
    onEnsureUser?: (login: string) => Promise<number>;
}

interface INeueTaState extends INewTaForm {
    saving: boolean;
    kundenList: string[];
    filteredKunden: string[];
    showKundenSuggestions: boolean;
    anwendungenList: string[];
    filteredAnwendungen: string[];
    showAnwendungenSuggestions: boolean;
    verantwortlicherSearchText: string;
    verantwortlicherLoginName: string | null;
    verantwortlicherSuggestions: any[];
    showVerantwortlicherSuggestions: boolean;
}

export default class NeueTa extends React.Component<INeueTaProps, INeueTaState> {
    constructor(props: INeueTaProps) {
        super(props);

        const kundenSet = new Set<string>();
        props.projekte.forEach(p => {
            // The user wants the Name (field_2) instead of the Number (field_1)
            if (p.field_2) kundenSet.add(p.field_2);
            // Fallback to field_1 if field_2 does not exist, though unlikely
            else if (p.field_1) kundenSet.add(p.field_1);
        });
        const kundenList = Array.from(kundenSet).sort();

        this.state = {
            kunde: '',
            endkunde: '',
            kontaktNr: '',
            aufgabe: '',
            anwendung: '',
            material: '',
            kategorie: '',
            potenzial: '',
            chance: '',
            budget: '',
            wunschtermin: '',
            verantwortlicherId: null,
            saving: false,
            kundenList,
            filteredKunden: [],
            showKundenSuggestions: false,
            anwendungenList: [],
            filteredAnwendungen: [],
            showAnwendungenSuggestions: false,
            verantwortlicherSearchText: '',
            verantwortlicherLoginName: null,
            verantwortlicherSuggestions: [],
            showVerantwortlicherSuggestions: false,
        };
    }

    private handleKundeChange = (value: string): void => {
        const filtered = this.state.kundenList.filter(k =>
            k.toLowerCase().includes(value.toLowerCase())
        );
        this.setState({
            kunde: value,
            filteredKunden: filtered,
            showKundenSuggestions: filtered.length > 0
        });

        // Auto-fill from project (find ALL projects that match the selected Name or Number)
        const projekteForKunde = this.props.projekte.filter(p =>
            (p.field_2 && p.field_2.trim() === value.trim()) ||
            (p.field_1 && p.field_1.trim() === value.trim())
        );

        // Populate "Anwendungen" for this specific customer
        const anwendungenSet = new Set<string>();
        projekteForKunde.forEach(p => { if (p.field_7) anwendungenSet.add(p.field_7); });

        const anwendungenList = Array.from(anwendungenSet).sort();

        if (projekteForKunde.length > 0) {
            this.setState({
                endkunde: projekteForKunde[0].field_13 || '',
                kontaktNr: projekteForKunde[0].Title || '',
                anwendungenList,
                filteredAnwendungen: [],
                showAnwendungenSuggestions: false,
            });
        } else {
            this.setState({
                anwendungenList: [],
                filteredAnwendungen: [],
                showAnwendungenSuggestions: false,
            });
        }
    }

    private selectKunde = (kunde: string): void => {
        this.handleKundeChange(kunde);
        this.setState({ showKundenSuggestions: false });
    }

    private handleAnwendungChange = (value: string): void => {
        const filtered = this.state.anwendungenList.filter(a =>
            a.toLowerCase().includes(value.toLowerCase())
        );
        this.setState({
            anwendung: value,
            filteredAnwendungen: filtered,
            showAnwendungenSuggestions: filtered.length > 0
        });

        // Auto fill specific application properties if found for the selected customer name and application
        const projekt = this.props.projekte.find(p =>
            ((p.field_2 && p.field_2.trim() === this.state.kunde.trim()) ||
                (p.field_1 && p.field_1.trim() === this.state.kunde.trim())) &&
            p.field_7 === value
        );
        if (projekt) {
            this.setState({
                // We don't overwrite material because we don't have it in IProjektItem currently, 
                // but we can auto-fill potential
                potenzial: projekt.field_8 ? projekt.field_8.toString() : this.state.potenzial,
                endkunde: projekt.field_13 || this.state.endkunde,
                kontaktNr: projekt.Title || this.state.kontaktNr,
                showAnwendungenSuggestions: false,
            });
        }
    }

    private selectAnwendung = (anwendung: string): void => {
        this.handleAnwendungChange(anwendung);
        this.setState({ showAnwendungenSuggestions: false });
    }

    private handleVerantwortlicherChange = async (value: string): Promise<void> => {
        this.setState({ verantwortlicherSearchText: value, verantwortlicherLoginName: null });
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

    private handleSubmit = async (): Promise<void> => {
        if (!this.state.kunde) return;
        this.setState({ saving: true });

        let finalVerantwortlicherId = this.state.verantwortlicherId;
        if (this.state.verantwortlicherLoginName && this.props.onEnsureUser) {
            try {
                finalVerantwortlicherId = await this.props.onEnsureUser(this.state.verantwortlicherLoginName);
            } catch (e) {
                console.error("User not found or ensure failed", e);
            }
        }

        try {
            await this.props.onSubmit({
                kunde: this.state.kunde,
                endkunde: this.state.endkunde,
                kontaktNr: this.state.kontaktNr,
                aufgabe: this.state.aufgabe,
                anwendung: this.state.anwendung,
                material: this.state.material,
                kategorie: this.state.kategorie,
                potenzial: this.state.potenzial,
                chance: this.state.chance,
                budget: this.state.budget,
                wunschtermin: this.state.wunschtermin,
                verantwortlicherId: finalVerantwortlicherId
            });
        } finally {
            this.setState({ saving: false });
        }
    }

    public render(): React.ReactElement<INeueTaProps> {
        const { saving } = this.state;

        return (
            <>
                {/* Header */}
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>Neue Technische Anfrage</span>
                    </div>
                </div>

                {/* Content */}
                <div className={styles.content}>
                    {/* Projekt-Infos */}
                    <div className={styles.formGroup} style={{ position: 'relative', zIndex: 10 }}>
                        <h4 className={styles.formGroupTitle}>Projekt-Infos {this.state.saving && '(Speichern...)'}</h4>

                        <div className={styles.formField} style={{ position: 'relative' }}>
                            <label className={styles.formLabel}>Kunde</label>
                            <input
                                className={styles.formInput}
                                value={this.state.kunde}
                                onChange={(e) => this.handleKundeChange(e.target.value)}
                                onFocus={() => this.setState({ showKundenSuggestions: true, filteredKunden: this.state.kunde ? this.state.filteredKunden : this.state.kundenList })}
                                onBlur={() => setTimeout(() => this.setState({ showKundenSuggestions: false }), 200)}
                                placeholder="Kunde suchen oder auswählen..."
                                style={{
                                    width: '100%',
                                    boxSizing: 'border-box',
                                    paddingRight: '32px',
                                    backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='rgba(0,0,0,0.5)' viewBox='0 0 16 16'%3E%3Cpath d='M8 11L3 6h10z'/%3E%3C/svg%3E")`,
                                    backgroundRepeat: 'no-repeat',
                                    backgroundPosition: 'right 12px center'
                                }}
                            />
                            {this.state.showKundenSuggestions && (
                                <div style={{ background: 'rgba(255,255,255,0.9)', border: '1px solid rgba(0,0,0,0.08)', borderRadius: 8, maxHeight: 150, overflow: 'auto', marginTop: 4 }}>
                                    {this.state.filteredKunden.slice(0, 50).map(k => (
                                        <div
                                            key={k}
                                            onClick={() => this.selectKunde(k)}
                                            style={{ padding: '8px 12px', cursor: 'pointer', fontSize: 13, color: 'rgb(15, 23, 42)' }}
                                            onMouseEnter={(e) => (e.currentTarget.style.background = 'rgba(0,0,0,0.04)')}
                                            onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                                        >
                                            {k}
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>

                        <div className={styles.formRow}>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Endkunde</label>
                                <input
                                    className={styles.formInput}
                                    value={this.state.endkunde}
                                    onChange={(e) => this.setState({ endkunde: e.target.value })}
                                />
                            </div>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Kontakt-Nr.</label>
                                <input
                                    className={styles.formInput}
                                    value={this.state.kontaktNr}
                                    readOnly
                                    style={{ opacity: 0.6 }}
                                />
                            </div>
                        </div>

                        <div className={styles.formField} style={{ position: 'relative' }}>
                            <label className={styles.formLabel}>Verantwortlicher</label>
                            <input
                                className={styles.formInput}
                                value={this.state.verantwortlicherSearchText}
                                onChange={(e) => this.handleVerantwortlicherChange(e.target.value)}
                                placeholder="Namen eingeben (min. 3 Zeichen)..."
                            />
                            {this.state.showVerantwortlicherSuggestions && (
                                <div style={{ position: 'absolute', zIndex: 10, background: 'rgb(255,255,255)', border: '1px solid rgba(0,0,0,0.1)', borderRadius: 4, maxHeight: 150, overflow: 'auto', width: '100%', top: 'calc(100% + 4px)', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}>
                                    {this.state.verantwortlicherSuggestions.map(u => (
                                        <div
                                            key={u.Key}
                                            onClick={() => this.selectVerantwortlicher(u)}
                                            style={{ padding: '8px 12px', cursor: 'pointer', fontSize: 13, color: 'rgb(15, 23, 42)' }}
                                            onMouseEnter={(e) => (e.currentTarget.style.background = 'rgba(0,0,0,0.04)')}
                                            onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                                        >
                                            {u.DisplayText}
                                            <div style={{ fontSize: 11, color: 'rgb(100, 116, 139)' }}>{u.EntityData?.Email || u.Description}</div>
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Anfrage-Details */}
                    <div className={styles.formGroup} style={{ position: 'relative', zIndex: 9 }}>
                        <h4 className={styles.formGroupTitle}>Anfrage-Details</h4>

                        <div className={styles.formField}>
                            <label className={styles.formLabel}>Aufgabenstellung</label>
                            <textarea
                                className={styles.formTextarea}
                                value={this.state.aufgabe}
                                onChange={(e) => this.setState({ aufgabe: e.target.value })}
                                placeholder="Beschreibe die technische Anfrage..."
                            />
                        </div>

                        <div className={styles.formRow}>
                            <div className={styles.formField} style={{ position: 'relative' }}>
                                <label className={styles.formLabel}>Anwendung</label>
                                <input
                                    className={styles.formInput}
                                    value={this.state.anwendung}
                                    onChange={(e) => this.handleAnwendungChange(e.target.value)}
                                    onFocus={() => this.setState({ showAnwendungenSuggestions: true, filteredAnwendungen: this.state.anwendung ? this.state.filteredAnwendungen : this.state.anwendungenList })}
                                    onBlur={() => setTimeout(() => this.setState({ showAnwendungenSuggestions: false }), 200)}
                                    placeholder={
                                        this.state.anwendungenList.length > 0
                                            ? "Anwendung aus Liste wählen..."
                                            : "Zuerst Kunde wählen..."
                                    }
                                    disabled={!this.state.kunde}
                                    style={{
                                        width: '100%',
                                        boxSizing: 'border-box',
                                        paddingRight: '32px',
                                        backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='rgba(0,0,0,0.5)' viewBox='0 0 16 16'%3E%3Cpath d='M8 11L3 6h10z'/%3E%3C/svg%3E")`,
                                        backgroundRepeat: 'no-repeat',
                                        backgroundPosition: 'right 12px center'
                                    }}
                                />
                                {this.state.showAnwendungenSuggestions && (
                                    <div style={{ background: 'rgba(255,255,255,0.9)', border: '1px solid rgba(0,0,0,0.08)', borderRadius: '0 0 8px 8px', maxHeight: 150, overflow: 'auto', marginTop: -2, position: 'relative', zIndex: 10 }}>
                                        {this.state.filteredAnwendungen.slice(0, 50).map(a => (
                                            <div
                                                key={a}
                                                onClick={() => this.selectAnwendung(a)}
                                                style={{ padding: '8px 12px', cursor: 'pointer', fontSize: 13, color: 'rgb(15, 23, 42)' }}
                                                onMouseEnter={(e) => (e.currentTarget.style.background = 'rgba(0,0,0,0.04)')}
                                                onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                                            >
                                                {a}
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Material</label>
                                <input
                                    className={styles.formInput}
                                    value={this.state.material}
                                    onChange={(e) => this.setState({ material: e.target.value })}
                                />
                            </div>
                        </div>

                        <div className={styles.formField}>
                            <label className={styles.formLabel}>Kategorie</label>
                            <select
                                className={styles.formSelect}
                                value={this.state.kategorie}
                                onChange={(e) => this.setState({ kategorie: e.target.value })}
                            >
                                <option value="">Auswählen...</option>
                                <option value="BM">BM</option>
                                <option value="PM">PM</option>
                                <option value="TA">TA</option>
                                <option value="MM">MM</option>
                            </select>
                        </div>
                    </div>

                    {/* Kaufmännisch */}
                    <div className={styles.formGroup} style={{ position: 'relative', zIndex: 8 }}>
                        <h4 className={styles.formGroupTitle}>Kaufmännisch & Termin</h4>

                        <div className={styles.formRow}>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Potenzial</label>
                                <div className={styles.formInputGroup}>
                                    <input
                                        type="number"
                                        className={styles.formInput}
                                        value={this.state.potenzial}
                                        onChange={(e) => this.setState({ potenzial: e.target.value })}
                                    />
                                    <span className={styles.formUnit}>jato</span>
                                </div>
                            </div>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Chance</label>
                                <div className={styles.formInputGroup}>
                                    <input
                                        type="number"
                                        className={styles.formInput}
                                        value={this.state.chance}
                                        onChange={(e) => this.setState({ chance: e.target.value })}
                                    />
                                    <span className={styles.formUnit}>%</span>
                                </div>
                            </div>
                            <div className={styles.formField}>
                                <label className={styles.formLabel}>Budget</label>
                                <div className={styles.formInputGroup}>
                                    <input
                                        type="number"
                                        className={styles.formInput}
                                        value={this.state.budget}
                                        onChange={(e) => this.setState({ budget: e.target.value })}
                                    />
                                    <span className={styles.formUnit}>€</span>
                                </div>
                            </div>
                        </div>
                    </div>

                    {/* Termine */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Termine</h4>
                        <div className={styles.formField}>
                            <label className={styles.formLabel}>Wunschtermin</label>
                            {(() => {
                                const wtParts = this.state.wunschtermin ? this.state.wunschtermin.split('.') : [];
                                const wtIso = wtParts.length === 3 ? `${wtParts[2]}-${wtParts[1]}-${wtParts[0]}` : '';
                                return (
                                    <input
                                        type="date"
                                        className={styles.dateInput}
                                        value={wtIso}
                                        onChange={(e) => {
                                            const parts = e.target.value.split('-');
                                            if (parts.length === 3) {
                                                this.setState({ wunschtermin: `${parts[2]}.${parts[1]}.${parts[0]}` });
                                            } else {
                                                this.setState({ wunschtermin: '' });
                                            }
                                        }}
                                    />
                                );
                            })()}
                        </div>
                    </div>
                </div>

                {/* Footer */}
                <div className={styles.btnRow}>
                    <button className={styles.btnSecondary} onClick={this.props.onBack}>
                        ✖ Abbrechen
                    </button>
                    <button
                        className={styles.btnPrimary}
                        onClick={this.handleSubmit}
                        disabled={!this.state.kunde || saving}
                    >
                        {saving ? '⏳ Speichere...' : '💾 TA anlegen'}
                    </button>
                </div>
            </>
        );
    }
}
