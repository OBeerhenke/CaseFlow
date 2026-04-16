import * as React from 'react';
import styles from './TaManagement.module.scss';
import { SharePointService } from '../services/SharePointService';
import { IKategorieItem, IKundenAnwendungItem } from '../models/types';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ISettingsProps {
    onBack: () => void;
    kundenList: string[]; // Pass this from parent to select a customer
}

interface ISettingsState {
    loading: boolean;
    error: string | null;

    // Kategorien
    kategorien: IKategorieItem[];
    newKategorieName: string;
    newKategorieEmail: string;
    editingEmailId: number | null;
    editingEmailValue: string;

    // View Navigation
    activeSettingView: 'menu' | 'kategorien' | 'config';

    // Global App Settings
    delayThreshold: string;
    savingConfig: boolean;
}

export default class Settings extends React.Component<ISettingsProps, ISettingsState> {
    constructor(props: ISettingsProps) {
        super(props);
        this.state = {
            loading: true,
            error: null,
            kategorien: [],
            newKategorieName: '',
            newKategorieEmail: '',
            editingEmailId: null,
            editingEmailValue: '',
            activeSettingView: 'menu',
            delayThreshold: '2',
            savingConfig: false
        };
    }

    public async componentDidMount(): Promise<void> {
        await this.loadKategorien();
    }

    private async loadKategorien(): Promise<void> {
        this.setState({ loading: true, error: null });
        try {
            const data = await SharePointService.instance.getKategorien();
            this.setState({ kategorien: data, loading: false });
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Laden der Kategorien: ' + e.message, loading: false });
        }
    }



    private handleAddKategorie = async (): Promise<void> => {
        if (!this.state.newKategorieName.trim()) return;
        this.setState({ loading: true });
        try {
            await SharePointService.instance.addKategorie(
                this.state.newKategorieName.trim(),
                this.state.newKategorieEmail.trim() || undefined
            );
            this.setState({ newKategorieName: '', newKategorieEmail: '' });
            await this.loadKategorien();
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Anlegen: ' + e.message, loading: false });
        }
    }

    private handleDeleteKategorie = async (id: number): Promise<void> => {
        if (!confirm('Soll diese Kategorie wirklich gelöscht werden?')) return;
        this.setState({ loading: true });
        try {
            await SharePointService.instance.deleteKategorie(id);
            await this.loadKategorien();
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Löschen: ' + e.message, loading: false });
        }
    }

    private handleSaveEmail = async (id: number): Promise<void> => {
        this.setState({ loading: true });
        try {
            await SharePointService.instance.updateKategorie(id, { Email: this.state.editingEmailValue.trim() || undefined });
            this.setState({ editingEmailId: null, editingEmailValue: '' });
            await this.loadKategorien();
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Speichern der Email: ' + e.message, loading: false });
        }
    }

    // ─── App Config ─────────────────────────────────────

    private async loadConfig(): Promise<void> {
        this.setState({ loading: true, error: null });
        try {
            const delayRaw = await SharePointService.instance.getConfigValue('DelayThresholdDays', '2');
            this.setState({ delayThreshold: delayRaw, loading: false });
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Laden der Einstellungen: ' + e.message, loading: false });
        }
    }

    private handleSaveConfig = async (): Promise<void> => {
        this.setState({ savingConfig: true, error: null });
        try {
            await SharePointService.instance.setConfigValue('DelayThresholdDays', this.state.delayThreshold);
            this.setState({ savingConfig: false });
            alert('Einstellungen gespeichert!');
        } catch (e: any) {
            this.setState({ error: 'Fehler beim Speichern der Einstellungen: ' + e.message, savingConfig: false });
        }
    }



    public render(): React.ReactElement<ISettingsProps> {
        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={() => {
                            if (this.state.activeSettingView !== 'menu') {
                                this.setState({ activeSettingView: 'menu' });
                            } else {
                                this.props.onBack();
                            }
                        }}>← Zurück</button>
                        <span className={styles.headerTitle}>
                            {this.state.activeSettingView === 'menu' ? 'Einstellungen & Hilfstabellen' :
                                this.state.activeSettingView === 'kategorien' ? 'TA Kategorien verwalten' :
                                    this.state.activeSettingView === 'config' ? 'Allgemeine App-Einstellungen' : ''}
                        </span>
                    </div>
                </div>

                {/* Main Settings Container - Takes up remaining height and prevents page scroll */}
                <div className={styles.content} style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden', padding: '16px' }}>
                    {this.state.error && (
                        <div style={{ width: '100%', padding: '12px', background: '#fee2e2', color: '#b91c1c', borderRadius: '6px', marginBottom: '1rem', flexShrink: 0 }}>
                            {this.state.error}
                        </div>
                    )}

                    {/* View: Menu */}
                    {this.state.activeSettingView === 'menu' && (
                        <div style={{ flex: 1, overflowY: 'auto' }}>
                            <h3 className={styles.sectionHeader}>Verfügbare Hilfstabellen</h3>

                            <div
                                className={styles.listRow}
                                style={{ alignItems: 'center', cursor: 'pointer' }}
                                onClick={() => this.setState({ activeSettingView: 'kategorien' })}
                            >
                                <span className={styles.listRowBold} style={{ width: '250px' }}>TA Kategorien</span>
                                <span className={styles.listRowMuted} style={{ flex: 1 }}>Auswahlmöglichkeiten für das Feld "Kategorie" bei Neuanlage.</span>
                                <button className={styles.btnPrimary} style={{ borderRadius: '20px', padding: '6px 16px', fontSize: '13px' }} onClick={(e) => { e.stopPropagation(); this.setState({ activeSettingView: 'kategorien' }); }}>
                                    Verwalten
                                </button>
                            </div>

                            <div
                                className={styles.listRow}
                                style={{ alignItems: 'center', cursor: 'pointer', marginTop: '12px' }}
                                onClick={() => { this.setState({ activeSettingView: 'config' }); void this.loadConfig(); }}
                            >
                                <span className={styles.listRowBold} style={{ width: '250px' }}>App-Einstellungen</span>
                                <span className={styles.listRowMuted} style={{ flex: 1 }}>Globale Systemvariablen (z.B. Schwellenwert für Verzögerungsgrund).</span>
                                <button className={styles.btnPrimary} style={{ borderRadius: '20px', padding: '6px 16px', fontSize: '13px' }} onClick={(e) => { e.stopPropagation(); this.setState({ activeSettingView: 'config' }); void this.loadConfig(); }}>
                                    Einstellungen
                                </button>
                            </div>

                        </div>
                    )}

                    {/* Flex Container for Column details - Scrollable internally if needed */}
                    {this.state.activeSettingView !== 'menu' && (
                        <div style={{ display: 'flex', gap: '24px', flex: 1, overflow: 'hidden' }}>

                            {/* Left Column: Kategorien */}
                            {this.state.activeSettingView === 'kategorien' && (
                                <div style={{ flex: '1 1 100%', display: 'flex', flexDirection: 'column', background: '#ffffff', borderRadius: '12px', padding: '24px', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', overflow: 'hidden' }}>
                                    <div style={{ flexShrink: 0 }}>
                                        <h3 style={{ marginTop: 0, borderBottom: '1px solid #e2e8f0', paddingBottom: '12px', color: '#1e293b' }}>
                                            TA Kategorien
                                        </h3>
                                        <p style={{ fontSize: '13px', color: '#64748b', marginBottom: '16px' }}>
                                            Auswahlmöglichkeiten für das Feld "Kategorie" bei Neuanlage.
                                        </p>

                                        <div style={{ display: 'flex', gap: '8px', marginBottom: '20px' }}>
                                            <input
                                                className={styles.formInput}
                                                style={{ flex: 1 }}
                                                placeholder="Neue Kategorie..."
                                                value={this.state.newKategorieName}
                                                onChange={(e) => this.setState({ newKategorieName: e.target.value })}
                                                onKeyPress={(e) => e.key === 'Enter' && this.handleAddKategorie()}
                                            />
                                            <input
                                                className={styles.formInput}
                                                style={{ flex: 1 }}
                                                placeholder="Email-Empfänger..."
                                                type="email"
                                                value={this.state.newKategorieEmail}
                                                onChange={(e) => this.setState({ newKategorieEmail: e.target.value })}
                                                onKeyPress={(e) => e.key === 'Enter' && this.handleAddKategorie()}
                                            />
                                            <button
                                                className={styles.btnPrimary}
                                                style={{ padding: '8px 16px', whiteSpace: 'nowrap' }}
                                                onClick={this.handleAddKategorie}
                                                disabled={!this.state.newKategorieName.trim() || this.state.loading}
                                            >
                                                Hinzufügen
                                            </button>
                                        </div>
                                    </div>

                                    {/* Scrollable List Container */}
                                    <div style={{ flex: 1, overflowY: 'auto', paddingRight: '4px' }}>
                                        {this.state.loading ? (
                                            <Spinner size={SpinnerSize.large} label="Lade Kategorien..." />
                                        ) : (
                                            <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                                                {this.state.kategorien.map(k => (
                                                    <li key={k.ID} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 12px', borderBottom: '1px solid #f1f5f9', gap: '12px' }}>
                                                        <span style={{ fontWeight: 500, color: '#334155', minWidth: '200px', flexShrink: 0 }}>{k.Title}</span>
                                                        <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                            {this.state.editingEmailId === k.ID ? (
                                                                <>
                                                                    <input
                                                                        className={styles.formInput}
                                                                        style={{ flex: 1, fontSize: '13px' }}
                                                                        type="email"
                                                                        placeholder="Email-Empfänger..."
                                                                        value={this.state.editingEmailValue}
                                                                        onChange={(e) => this.setState({ editingEmailValue: e.target.value })}
                                                                        onKeyPress={(e) => e.key === 'Enter' && this.handleSaveEmail(k.ID)}
                                                                        autoFocus
                                                                    />
                                                                    <button
                                                                        style={{ background: 'none', border: 'none', color: '#10b981', cursor: 'pointer', fontSize: '14px', padding: '2px 4px' }}
                                                                        onClick={() => this.handleSaveEmail(k.ID)}
                                                                        title="Speichern"
                                                                    >✓</button>
                                                                    <button
                                                                        style={{ background: 'none', border: 'none', color: '#94a3b8', cursor: 'pointer', fontSize: '14px', padding: '2px 4px' }}
                                                                        onClick={() => this.setState({ editingEmailId: null, editingEmailValue: '' })}
                                                                        title="Abbrechen"
                                                                    >✕</button>
                                                                </>
                                                            ) : (
                                                                <span
                                                                    style={{ fontSize: '13px', color: k.Email ? '#3b82f6' : '#94a3b8', cursor: 'pointer' }}
                                                                    onClick={() => this.setState({ editingEmailId: k.ID, editingEmailValue: k.Email || '' })}
                                                                    title="Email bearbeiten"
                                                                >
                                                                    {k.Email || '(keine Email)'}
                                                                </span>
                                                            )}
                                                        </div>
                                                        <button
                                                            style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '16px', flexShrink: 0 }}
                                                            onClick={() => this.handleDeleteKategorie(k.ID)}
                                                            title="Löschen"
                                                        >
                                                            <Icon iconName="Delete" />
                                                        </button>
                                                    </li>
                                                ))}
                                                {this.state.kategorien.length === 0 && (
                                                    <li style={{ color: '#94a3b8', fontSize: '13px', textAlign: 'center', padding: '20px' }}>Keine Kategorien vorhanden.</li>
                                                )}
                                            </ul>
                                        )}
                                    </div>
                                </div>
                            )}

                            {/* ───────────────────────────────────────────────────────── */}
                            {/* CONFIG VIEW                                           */}
                            {/* ───────────────────────────────────────────────────────── */}
                            {this.state.activeSettingView === 'config' && (
                                <div style={{ flex: '1 1 100%', display: 'flex', flexDirection: 'column', background: '#ffffff', borderRadius: '12px', padding: '24px', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)', overflow: 'hidden' }}>
                                    <h3 style={{ marginTop: 0, borderBottom: '1px solid #e2e8f0', paddingBottom: '12px', color: '#1e293b' }}>
                                        Allgemeine App-Einstellungen
                                    </h3>

                                    <div style={{ maxWidth: 600, marginTop: '16px' }}>
                                        <div className={styles.formField} style={{ marginBottom: 24 }}>
                                            <label className={styles.formLabel}>Verzögerungsgrund (Tage)</label>
                                            <p style={{ fontSize: 13, color: '#64748b', margin: '4px 0 12px 0' }}>
                                                Legt fest, ab wie vielen Tagen (zwischen Erstellungsdatum und Planungsdatum) der User gezwungen wird, einen Verzögerungsgrund beim Planen anzugeben.
                                            </p>
                                            <input
                                                type="number"
                                                className={styles.formInput}
                                                value={this.state.delayThreshold}
                                                onChange={(e) => this.setState({ delayThreshold: e.target.value })}
                                                style={{ width: '150px' }}
                                                min="0"
                                            />
                                        </div>

                                        <button
                                            className={styles.btnPrimary}
                                            onClick={this.handleSaveConfig}
                                            disabled={this.state.savingConfig}
                                        >
                                            {this.state.savingConfig ? 'Speichere...' : 'Einstellungen speichern'}
                                        </button>
                                    </div>
                                </div>
                            )}

                        </div>
                    )}
                </div>
            </>
        );
    }
}
