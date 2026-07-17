import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, STATUS_VALUES, INITIAL_DELAY_REASONS } from '../models/types';
import { SharePointService } from '../services/SharePointService';
import StatusPill from './StatusPill';
import Modal from './Modal';

/** Format ISO date string (e.g. "2026-04-20T07:00:00Z") to German locale "20.04.2026" */
function formatDate(value: string | undefined | null): string {
    if (!value) return '–';
    const d = new Date(value);
    if (isNaN(d.getTime())) return value;
    return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

export interface ITaDetailProps {
    ta: ITaItem;
    onBack: () => void;
    onSave: (id: number, fields: Partial<Record<string, string | number | undefined>>) => Promise<void>;
    onVerschieben: (id: number, neuerTermin: string, grund: string, alterTermin: string, urspruenglicherTermin?: string) => Promise<void>;
    onSetTermin?: (id: number, termin: string, verantwortlicherId?: number, delayReason?: string) => Promise<void>;
    onSearchUsers?: (query: string) => Promise<any[]>;
    onEnsureUser?: (login: string) => Promise<number>;
    currentUserDisplayName: string;
    currentUserLoginName: string;
}

interface ITaDetailState {
    bemerkungHistorie: string;
    neueBemerkung: string;
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
    // Editable fields
    prioritaet: string;
    budgetBeiStart: string;
    istKosten: string;
    geplanteKosten: string;
    // Attachments
    attachments: { FileName: string; ServerRelativeUrl: string }[];
    uploadingFile: boolean;
}

interface IBemerkungEntry {
    timestamp: string;
    kuerzel: string;
    text: string;
}

export default class TaDetail extends React.Component<ITaDetailProps, ITaDetailState> {
    private fileInputRef = React.createRef<HTMLInputElement>();

    private normalizeIsoDateInput(value: string): string {
        if (!value) return '';
        const match = value.match(/^(\d{0,4})(?:-(\d{0,2}))?(?:-(\d{0,2}))?/);
        if (!match) return '';

        const year = (match[1] || '').slice(0, 4);
        const month = (match[2] || '').slice(0, 2);
        const day = (match[3] || '').slice(0, 2);

        if (!month && !day) return year;
        if (!day) return `${year}-${month}`;
        return `${year}-${month}-${day}`;
    }

    constructor(props: ITaDetailProps) {
        super(props);
        this.state = {
            bemerkungHistorie: props.ta.field_2 || '',
            neueBemerkung: '',
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
            delayThresholdDays: 2,
            // Editable fields init from TA
            prioritaet: props.ta.field_13 !== undefined && props.ta.field_13 !== null ? props.ta.field_13.toString() : '',
            budgetBeiStart: props.ta.field_17 || '',
            istKosten: props.ta.field_18 !== undefined && props.ta.field_18 !== null ? props.ta.field_18.toString() : '',
            geplanteKosten: props.ta.field_19 !== undefined && props.ta.field_19 !== null ? props.ta.field_19.toString() : '',
            attachments: [],
            uploadingFile: false
        };
    }

    public async componentDidMount(): Promise<void> {
        try {
            const delayRaw = await SharePointService.instance.getConfigValue('DelayThresholdDays', '2');
            this.setState({ delayThresholdDays: parseInt(delayRaw, 10) || 2 });
        } catch (e) {
            console.error("Failed to load delay threshold config", e);
        }
        this.loadAttachments().catch(console.error);
    }

    private async loadAttachments(): Promise<void> {
        try {
            const attachments = await SharePointService.instance.getAttachments(this.props.ta.ID);
            this.setState({ attachments });
        } catch (e) {
            console.error("Failed to load attachments", e);
        }
    }

    private handleFileUpload = async (files: FileList): Promise<void> => {
        this.setState({ uploadingFile: true });
        try {
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const buffer = await file.arrayBuffer();
                await SharePointService.instance.addAttachment(this.props.ta.ID, file.name, buffer);
            }
            await this.loadAttachments();
        } catch (e) {
            console.error("Failed to upload attachment", e);
        } finally {
            this.setState({ uploadingFile: false });
        }
    }

    private handleDeleteAttachment = async (fileName: string): Promise<void> => {
        try {
            await SharePointService.instance.deleteAttachment(this.props.ta.ID, fileName);
            await this.loadAttachments();
        } catch (e) {
            console.error("Failed to delete attachment", e);
        }
    }

    private getUserKuerzel(): string {
        return this.getAccountKuerzel(this.props.currentUserLoginName, this.props.currentUserDisplayName);
    }

    private getAccountKuerzel(emailOrLogin?: string, fallbackName?: string): string {
        const raw = (emailOrLogin || '').trim();
        if (raw) {
            const claimPart = raw.includes('|') ? (raw.split('|').pop() || raw) : raw;
            const slashPart = claimPart.includes('\\') ? (claimPart.split('\\').pop() || claimPart) : claimPart;
            const localPart = slashPart.includes('@') ? slashPart.split('@')[0] : slashPart;
            const normalized = localPart.replace(/[^a-zA-Z0-9]/g, '');
            if (normalized) return normalized.substring(0, 6).toUpperCase();
        }

        if (fallbackName) {
            const parts = fallbackName.trim().split(/\s+/).filter(Boolean);
            if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
            if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
        }

        return 'NA';
    }

    private getLegacyTimestamp(): string {
        const created = this.props.ta.field_4;
        if (!created) return this.formatBemerkungTimestamp(new Date());

        const createdDate = new Date(created);
        if (!isNaN(createdDate.getTime())) return this.formatBemerkungTimestamp(createdDate);

        const parts = created.split('.');
        if (parts.length === 3) {
            const fallback = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T00:00:00`);
            if (!isNaN(fallback.getTime())) return this.formatBemerkungTimestamp(fallback);
        }

        return this.formatBemerkungTimestamp(new Date());
    }

    private formatBemerkungTimestamp(date: Date): string {
        const d = date.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const t = date.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
        return `${d} ${t}`;
    }

    private parseBemerkungEntries(history: string): IBemerkungEntry[] {
        const lines = (history || '').split(/\r?\n/);
        const entries: IBemerkungEntry[] = [];
        const headerRegex = /^\[(\d{2}\.\d{2}\.\d{4}\s\d{2}:\d{2})\]\s\[([^\]]+)\]\s?(.*)$/;

        let current: IBemerkungEntry | null = null;
        const legacyTimestamp = this.getLegacyTimestamp();
        const legacyKuerzel = this.getAccountKuerzel(this.props.ta.Ersteller?.EMail, this.props.ta.Ersteller?.Title);

        const pushCurrent = (): void => {
            if (!current) return;
            const text = current.text.trim();
            if (text) entries.push({ ...current, text });
            current = null;
        };

        for (const line of lines) {
            const headerMatch = line.match(headerRegex);
            if (headerMatch) {
                pushCurrent();
                current = {
                    timestamp: headerMatch[1],
                    kuerzel: headerMatch[2],
                    text: headerMatch[3] || ''
                };
                continue;
            }

            if (!current) {
                if (!line.trim()) continue;
                current = {
                    timestamp: legacyTimestamp,
                    kuerzel: legacyKuerzel,
                    text: line
                };
                continue;
            }

            current.text = current.text ? `${current.text}\n${line}` : line;
        }

        pushCurrent();
        return this.splitLegacyMixedEntries(entries);
    }

    private splitLegacyMixedEntries(entries: IBemerkungEntry[]): IBemerkungEntry[] {
        const result: IBemerkungEntry[] = [];
        const legacyTimestamp = this.getLegacyTimestamp();
        const legacyKuerzel = this.getAccountKuerzel(this.props.ta.Ersteller?.EMail, this.props.ta.Ersteller?.Title);

        for (const entry of entries) {
            const lines = entry.text
                .split(/\r?\n/)
                .map(l => l.trim())
                .filter(Boolean);

            if (lines.length <= 1) {
                result.push({ ...entry, text: entry.text.trim() });
                continue;
            }

            // Legacy migration: previously appended comments could contain
            // one new headered line plus older plain-text lines.
            result.push({ ...entry, text: lines[0] });

            for (let i = 1; i < lines.length; i++) {
                result.push({
                    timestamp: legacyTimestamp,
                    kuerzel: legacyKuerzel,
                    text: lines[i]
                });
            }
        }

        return result;
    }

    private serializeBemerkungEntries(entries: IBemerkungEntry[]): string {
        return entries
            .filter(e => e.text.trim())
            .map(e => `[${e.timestamp}] [${e.kuerzel}] ${e.text.trim()}`)
            .join('\n\n');
    }

    private buildBemerkungHistoryPayload(): { nextHistory: string; changed: boolean } {
        const existingEntries = this.parseBemerkungEntries(this.state.bemerkungHistorie);
        const normalizedExisting = this.serializeBemerkungEntries(existingEntries);

        const newText = this.state.neueBemerkung.trim();
        let nextHistory = normalizedExisting;

        if (newText) {
            const newEntry: IBemerkungEntry = {
                timestamp: this.formatBemerkungTimestamp(new Date()),
                kuerzel: this.getUserKuerzel(),
                text: newText
            };
            nextHistory = this.serializeBemerkungEntries([newEntry, ...existingEntries]);
        }

        return {
            nextHistory,
            changed: nextHistory !== this.state.bemerkungHistorie
        };
    }

    private handleSave = async (): Promise<void> => {
        this.setState({ saving: true });
        try {
            const updates: Partial<Record<string, string | number | undefined>> = {
                Status: this.state.status,
                // Save priority + cost fields
                field_13: this.state.prioritaet ? parseInt(this.state.prioritaet, 10) : undefined,
                field_17: this.state.budgetBeiStart || undefined,
                field_18: this.state.istKosten ? parseFloat(this.state.istKosten.replace(',', '.')) : undefined,
                field_19: this.state.geplanteKosten ? parseFloat(this.state.geplanteKosten.replace(',', '.')) : undefined,
            };

            const bemerkungPayload = this.buildBemerkungHistoryPayload();
            if (bemerkungPayload.changed) {
                updates.field_2 = bemerkungPayload.nextHistory;
            }

            // If completing the TA, set Erledigungsdatum to today
            if (this.state.status === 'abgeschlossen' && this.props.ta.Status !== 'abgeschlossen') {
                const today = new Date();
                const pad2 = (n: number): string => n < 10 ? '0' + n : '' + n;
                updates.Erledigungsdatum = today.getFullYear() + '-' + pad2(today.getMonth() + 1) + '-' + pad2(today.getDate());
            }

            // If Verantwortlicher was changed, resolve and save
            if (this.state.verantwortlicherLoginName && this.props.onEnsureUser) {
                try {
                    const userId = await this.props.onEnsureUser(this.state.verantwortlicherLoginName);
                    (updates as any).VerantwortlicherId = userId;
                } catch (e) {
                    console.error("Error ensuring user for save:", e);
                }
            }

            await this.props.onSave(this.props.ta.ID, updates);

            if (bemerkungPayload.changed) {
                this.setState({ bemerkungHistorie: bemerkungPayload.nextHistory, neueBemerkung: '' });
            }
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
            // Save termin + verantwortlicher
            await this.props.onSetTermin(this.props.ta.ID, formatted, finalVerantwortlicherId, this.state.delayReason);

            // Also save costs + bemerkung in the same action
            const updates: Partial<Record<string, string | number | undefined>> = {
                field_17: this.state.budgetBeiStart || undefined,
                field_18: this.state.istKosten ? parseFloat(this.state.istKosten.replace(',', '.')) : undefined,
                field_19: this.state.geplanteKosten ? parseFloat(this.state.geplanteKosten.replace(',', '.')) : undefined,
            };

            const bemerkungPayload = this.buildBemerkungHistoryPayload();
            if (bemerkungPayload.changed) {
                updates.field_2 = bemerkungPayload.nextHistory;
            }

            await this.props.onSave(this.props.ta.ID, updates);

            if (bemerkungPayload.changed) {
                this.setState({ bemerkungHistorie: bemerkungPayload.nextHistory, neueBemerkung: '' });
            }
        } finally {
            this.setState({ savingTermin: false, verantwortlicherSearchText: '', verantwortlicherLoginName: null, termin: '', showDelayModal: false, delayReason: '', neueBemerkung: '' });
        }
    }

    private handleVerschieben = async (neuerTermin: string, grund: string): Promise<void> => {
        if (this.state.saving) return;
        this.setState({ showModal: false });
        if (!window.confirm(`Termin wirklich auf ${neuerTermin} verschieben?\nGrund: ${grund}`)) return;
        this.setState({ saving: true });
        try {
            await this.props.onVerschieben(
                this.props.ta.ID,
                neuerTermin,
                grund,
                this.props.ta.field_6 || '',
                this.props.ta.field_22 || undefined
            );
        } finally {
            this.setState({ saving: false });
        }
    }

    private formatCurrency(val: number | undefined): string {
        if (val === undefined || val === null) return '–';
        return val.toLocaleString('de-DE') + ' €';
    }

    public render(): React.ReactElement<ITaDetailProps> {
        const { ta } = this.props;
        const bemerkungEntries = this.parseBemerkungEntries(this.state.bemerkungHistorie);

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
                                Erstellt: {formatDate(ta.field_4)}
                            </span>
                        </div>
                    </div>

                    {/* Termine */}
                    <div className={styles.formGroup} style={{ position: 'relative', zIndex: 10 }}>
                        <h4 className={styles.formGroupTitle}>Termine</h4>
                        <div className={styles.detailGrid}>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Wunschtermin</span>
                                <span className={styles.detailValue}>{formatDate(ta.field_5)}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Geplanter Termin</span>
                                <span className={styles.detailValue}>{formatDate(ta.field_6)}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Ursprünglicher Termin</span>
                                <span className={styles.detailValue}>{formatDate(ta.field_22)}</span>
                            </div>
                        </div>

                        {!ta.field_6 ? (
                            <div style={{ marginTop: '16px', padding: '16px', background: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                                <h5 style={{ margin: '0 0 12px 0', color: '#334155' }}>Neuen Termin einplanen</h5>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>

                                    {/* Verantwortlicher Search */}
                                    <div style={{ position: 'relative' }}>
                                        <label className={styles.formLabel} style={{ fontSize: '12px', marginBottom: '4px' }}>Verantwortlicher <span style={{ color: '#ef4444' }}>*</span></label>
                                        <input
                                            className={styles.formInput}
                                            style={{ width: '100%' }}
                                            value={this.state.verantwortlicherSearchText}
                                            onChange={(e) => this.handleVerantwortlicherChange(e.target.value)}
                                            placeholder="Suchen (min. 3 Zeichen)"
                                        />
                                        {this.state.showVerantwortlicherSuggestions && (
                                            <div className={styles.suggestionDropdown}>
                                                {this.state.verantwortlicherSuggestions.map(u => (
                                                    <div
                                                        key={u.Key}
                                                        className={styles.suggestionItem}
                                                        onClick={() => this.selectVerantwortlicher(u)}
                                                    >
                                                        <div className={styles.suggestionAvatar}>
                                                            {(u.DisplayText || '?').charAt(0).toUpperCase()}
                                                        </div>
                                                        <div className={styles.suggestionInfo}>
                                                            <span className={styles.suggestionName}>{u.DisplayText}</span>
                                                            <span className={styles.suggestionEmail}>{u.EntityData?.Email || u.Description || ''}</span>
                                                        </div>
                                                    </div>
                                                ))}
                                            </div>
                                        )}
                                    </div>

                                    {/* Date Picker */}
                                    <div>
                                        <label className={styles.formLabel} style={{ fontSize: '12px', marginBottom: '4px' }}>Datum</label>
                                        <input
                                            type="date"
                                            className={styles.formInput}
                                            style={{ width: '100%' }}
                                            value={this.state.termin}
                                            onChange={(e) => {
                                                const normalized = this.normalizeIsoDateInput(e.target.value);
                                                if (normalized !== e.target.value) {
                                                    e.currentTarget.value = normalized;
                                                }
                                                this.setState({ termin: normalized });
                                            }}
                                        />
                                    </div>
                                </div>
                            </div>
                        ) : (
                            <div style={{ marginTop: '16px', display: 'flex', gap: '12px', alignItems: 'flex-end', flexWrap: 'wrap' }}>
                                {/* Verantwortlicher ändern */}
                                <div style={{ flex: '1 1 200px', position: 'relative' }}>
                                    <label className={styles.formLabel} style={{ fontSize: '12px', marginBottom: '4px' }}>Verantwortlicher ändern</label>
                                    <input
                                        className={styles.formInput}
                                        style={{ width: '100%' }}
                                        value={this.state.verantwortlicherSearchText}
                                        onChange={(e) => this.handleVerantwortlicherChange(e.target.value)}
                                        placeholder={ta.Verantwortlicher?.Title || 'Suchen (min. 3 Zeichen)'}
                                    />
                                    {this.state.showVerantwortlicherSuggestions && (
                                        <div className={styles.suggestionDropdown}>
                                            {this.state.verantwortlicherSuggestions.map(u => (
                                                <div
                                                    key={u.Key}
                                                    className={styles.suggestionItem}
                                                    onClick={() => this.selectVerantwortlicher(u)}
                                                >
                                                    <div className={styles.suggestionAvatar}>
                                                        {(u.DisplayText || '?').charAt(0).toUpperCase()}
                                                    </div>
                                                    <div className={styles.suggestionInfo}>
                                                        <span className={styles.suggestionName}>{u.DisplayText}</span>
                                                        <span className={styles.suggestionEmail}>{u.EntityData?.Email || u.Description || ''}</span>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </div>
                                <button
                                    className={styles.btnWarning}
                                    onClick={() => this.setState({ showModal: true })}
                                >
                                    Termin verschieben
                                </button>
                            </div>
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
                                <span className={styles.detailLabel}>Kontaktnummer</span>
                                <span className={styles.detailValue}>{ta.field_10 || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>VP</span>
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
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>SOP</span>
                                <span className={styles.detailValue}>{ta.SOP || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>SegCode</span>
                                <span className={styles.detailValue}>{ta.SegCode || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Antwort in</span>
                                <span className={styles.detailValue}>{ta.AntwortIn || '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Zielpreis</span>
                                <span className={styles.detailValue}>{ta.Zielpreis != null ? ta.Zielpreis.toLocaleString('de-DE') + ' €' : '–'}</span>
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Priorität</span>
                                <select
                                    className={styles.formSelect}
                                    value={this.state.prioritaet}
                                    onChange={(e) => this.setState({ prioritaet: e.target.value })}
                                >
                                    <option value="">Ohne</option>
                                    <option value="1">▼ Niedrig</option>
                                    <option value="2">● Mittel</option>
                                    <option value="3">▲ Hoch</option>
                                </select>
                            </div>
                        </div>
                    </div>

                    {/* Kosten */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Kosten</h4>
                        <div className={styles.detailGrid}>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Budget bei Start</span>
                                <input
                                    className={styles.formInput}
                                    value={this.state.budgetBeiStart}
                                    onChange={(e) => this.setState({ budgetBeiStart: e.target.value })}
                                    placeholder="z.B. 5.000 €"
                                    style={{ color: !this.state.budgetBeiStart || parseFloat(this.state.budgetBeiStart) === 0 ? '#EF4444' : '#3B82F6', fontWeight: 600 }}
                                />
                                {(!this.state.budgetBeiStart || parseFloat(this.state.budgetBeiStart) === 0) && (
                                    <span style={{ color: '#EF4444', fontSize: 11, marginTop: 4, display: 'block' }}>⚠ Kein Projektbudget vorhanden</span>
                                )}
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>Geplante Kosten (€) {this.state.status === 'Termin planen' && <span style={{ color: '#ef4444' }}>*</span>}</span>
                                {this.state.status === 'Termin planen' ? (
                                    <input
                                        type="number"
                                        step="0.01"
                                        className={styles.formInput}
                                        value={this.state.geplanteKosten}
                                        onChange={(e) => this.setState({ geplanteKosten: e.target.value })}
                                        placeholder="0.00"
                                    />
                                ) : (
                                    <span className={styles.detailValue}>
                                        {this.state.geplanteKosten ? `${parseFloat(this.state.geplanteKosten).toLocaleString('de-DE')} €` : '–'}
                                    </span>
                                )}
                            </div>
                            <div className={styles.detailItem}>
                                <span className={styles.detailLabel}>IST Kosten (€)</span>
                                {this.state.status !== 'Termin planen' ? (
                                    <input
                                        type="number"
                                        step="0.01"
                                        className={styles.formInput}
                                        value={this.state.istKosten}
                                        onChange={(e) => this.setState({ istKosten: e.target.value })}
                                        placeholder="0.00"
                                    />
                                ) : (
                                    <span className={styles.detailValue}>
                                        {this.state.istKosten ? `${parseFloat(this.state.istKosten).toLocaleString('de-DE')} €` : '–'}
                                    </span>
                                )}
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

                    {/* Aufgabenstellung */}
                    {ta.Aufgabenstellung && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>Aufgabenstellung</h4>
                            <p style={{ fontSize: 13, color: 'rgb(15, 23, 42)', margin: 0, whiteSpace: 'pre-wrap' }}>{ta.Aufgabenstellung}</p>
                        </div>
                    )}

                    {/* Bemerkungen */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Bemerkungen</h4>
                        <span className={styles.detailLabel}>Historie</span>
                        {bemerkungEntries.length > 0 ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 10 }}>
                                {bemerkungEntries.map((entry, index) => (
                                    <div
                                        key={`${entry.timestamp}-${entry.kuerzel}-${index}`}
                                        style={{
                                            border: '1px solid rgba(148, 163, 184, 0.35)',
                                            borderRadius: 10,
                                            padding: '10px 12px',
                                            background: 'rgba(255, 255, 255, 0.5)'
                                        }}
                                    >
                                        <div style={{ fontSize: 12, color: '#64748b', marginBottom: 4, fontWeight: 600 }}>
                                            {entry.timestamp} | {entry.kuerzel}
                                        </div>
                                        <div style={{ fontSize: 14, color: '#0f172a', whiteSpace: 'pre-wrap' }}>
                                            {entry.text}
                                        </div>
                                    </div>
                                ))}
                            </div>
                        ) : (
                            <p style={{ margin: '6px 0 10px 0', color: '#64748b', fontSize: 13 }}>Keine Bemerkungen vorhanden</p>
                        )}
                        <span className={styles.detailLabel}>Neue Bemerkung hinzufügen</span>
                        <textarea
                            className={styles.formTextarea}
                            value={this.state.neueBemerkung}
                            onChange={(e) => this.setState({ neueBemerkung: e.target.value })}
                            placeholder="Neue Bemerkung eingeben. Beim Speichern wird sie mit Datum/Uhrzeit und Kürzel oben angefügt."
                        />
                    </div>

                    {/* Anhänge */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Anhänge</h4>
                        <input
                            ref={this.fileInputRef}
                            type="file"
                            multiple
                            style={{ display: 'none' }}
                            onChange={(e) => {
                                if (e.target.files && e.target.files.length > 0) {
                                    this.handleFileUpload(e.target.files).catch(console.error);
                                }
                                e.target.value = '';
                            }}
                        />
                        <button
                            type="button"
                            className={styles.btnSecondary}
                            onClick={() => this.fileInputRef.current?.click()}
                            disabled={this.state.uploadingFile}
                            style={{ marginBottom: 8 }}
                        >
                            {this.state.uploadingFile ? 'Wird hochgeladen...' : '+ Dateien hinzufügen'}
                        </button>
                        {this.state.attachments.length > 0 ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                {this.state.attachments.map(att => (
                                    <div key={att.FileName} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13 }}>
                                        <a
                                            href={att.ServerRelativeUrl}
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            style={{ color: '#3b82f6', textDecoration: 'none' }}
                                        >
                                            📎 {att.FileName}
                                        </a>
                                        <button
                                            type="button"
                                            onClick={() => { if (confirm(`"${att.FileName}" wirklich löschen?`)) this.handleDeleteAttachment(att.FileName).catch(console.error); }}
                                            style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: 13, padding: 0 }}
                                        >
                                            ✕
                                        </button>
                                    </div>
                                ))}
                            </div>
                        ) : (
                            <p style={{ fontSize: 13, color: '#94a3b8', margin: 0 }}>Keine Anhänge</p>
                        )}
                    </div>

                    {/* Verschiebungsgrund */}
                    {ta.field_21 && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>Letzter Verschiebungsgrund</h4>
                            <p style={{ fontSize: 13, color: '#F59E0B', margin: 0 }}>{ta.field_21}</p>
                        </div>
                    )}

                    {/* Aktionen */}
                    {this.state.status !== 'abgeschlossen' && this.state.status !== 'Termin planen' && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>Aktionen</h4>
                            <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                                <button
                                    type="button"
                                    className={styles.btnPrimary}
                                    onClick={this.handleSave}
                                    disabled={this.state.saving}
                                >
                                    {this.state.saving ? 'Speichere...' : 'Änderungen speichern'}
                                </button>
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
                                    disabled={this.state.saving || !this.state.istKosten}
                                >
                                    TA erfolgreich abschließen
                                </button>
                            </div>
                            {!this.state.istKosten && (
                                <p style={{ fontSize: 12, color: '#ef4444', margin: '8px 0 0 0' }}>Bitte IST Kosten eintragen, um die TA abzuschließen.</p>
                            )}
                        </div>
                    )}
                </div>

                {/* Footer */}
                <div className={styles.btnRow}>
                    <button className={styles.btnSecondary} onClick={this.props.onBack}>
                        ← Zurück
                    </button>
                    {this.state.status === 'Termin planen' && (
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
                            <button
                                className={styles.btnSuccess}
                                onClick={this.handlePlanen}
                                disabled={!this.state.termin || !this.state.verantwortlicherLoginName || !this.state.geplanteKosten || this.state.savingTermin}
                            >
                                {this.state.savingTermin ? 'Speichere...' : 'Termin einplanen & Speichern'}
                            </button>
                            {(!this.state.termin || !this.state.verantwortlicherLoginName || !this.state.geplanteKosten) && (
                                <span style={{ fontSize: 12, color: '#ef4444' }}>
                                    Pflichtfelder: {[!this.state.termin && 'Termin', !this.state.verantwortlicherLoginName && 'Verantwortlicher', !this.state.geplanteKosten && 'Geplante Kosten'].filter(Boolean).join(', ')}
                                </span>
                            )}
                        </div>
                    )}
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
