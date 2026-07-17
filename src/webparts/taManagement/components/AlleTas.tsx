import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, STATUS_VALUES } from '../models/types';
import StatusPill from './StatusPill'; 

type SortKey = 'Title' | 'Ersteller' | 'Kunde' | 'Material' | 'Kategorie' | 'Verantwortlicher' | 'Termin' | 'Status';
type SortDir = 'asc' | 'desc';

const getAccountKuerzel = (emailOrLogin?: string, fallbackName?: string): string => {
    const raw = (emailOrLogin || '').trim();
    if (raw) {
        const claimPart = raw.includes('|') ? (raw.split('|').pop() || raw) : raw;
        const slashPart = claimPart.includes('\\') ? (claimPart.split('\\').pop() || claimPart) : claimPart;
        const localPart = slashPart.includes('@') ? slashPart.split('@')[0] : slashPart;
        const tokens = localPart.split(/[^a-zA-Z0-9]+/).filter(Boolean);
        if (tokens.length >= 2) return (tokens[0][0] + tokens[tokens.length - 1][0]).toUpperCase();

        const normalized = localPart.replace(/[^a-zA-Z0-9]/g, '');
        if (normalized.length > 0 && normalized.length <= 4) return normalized.toUpperCase();
        if (normalized.length >= 2) return normalized.substring(0, 2).toUpperCase();
    }

    if (fallbackName) {
        const parts = fallbackName.trim().split(/\s+/).filter(Boolean);
        if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    }

    return '–';
};

const getKuerzel = (u?: { Title?: string; EMail?: string; Name?: string; Nickname?: string }): string => {
    const nick = (u?.Nickname || '').trim();
    if (nick) return nick.toUpperCase();
    return getAccountKuerzel(u?.Name || u?.EMail, u?.Title);
};

export interface IAlleTasProps {
    tas: ITaItem[];
    initialFilter: string;
    onSelectTa: (id: number) => void;
    onBack: () => void;
}

interface IAlleTasState {
    search: string;
    statusFilter: string;
    kategorieFilter: string;
    sortKey: SortKey;
    sortDir: SortDir;
}

export default class AlleTas extends React.Component<IAlleTasProps, IAlleTasState> {
    constructor(props: IAlleTasProps) {
        super(props);
        this.state = {
            search: '',
            statusFilter: props.initialFilter || 'Offen',
            kategorieFilter: 'Alle',
            sortKey: 'Termin',
            sortDir: 'asc'
        };
    }

    public componentDidUpdate(prevProps: IAlleTasProps): void {
        if (prevProps.initialFilter !== this.props.initialFilter && this.props.initialFilter) {
            this.setState({ statusFilter: this.props.initialFilter });
        }
    }

    private getKategorien(): string[] {
        const set = new Set<string>();
        this.props.tas.forEach(ta => { if (ta.field_16) set.add(ta.field_16); });
        return Array.from(set).sort();
    }

    private getSortValue(ta: ITaItem, key: SortKey): string {
        switch (key) {
            case 'Title': return ta.Title || '';
            case 'Ersteller': return ta.Ersteller?.Title || '';
            case 'Kunde': return ta.field_8 || '';
            case 'Material': return ta.field_12 || '';
            case 'Kategorie': return ta.field_16 || '';
            case 'Verantwortlicher': return ta.Verantwortlicher?.Title || '';
            case 'Termin': return ta.field_6 || '9999-12-31';
            case 'Status': return ta.Status || '';
            default: return '';
        }
    }

    private getFilteredTas(): ITaItem[] {
        const { search, statusFilter, kategorieFilter, sortKey, sortDir } = this.state;

        const filtered = this.props.tas.filter(ta => {
            // Text search
            if (search) {
                const s = search.toLowerCase();
                const ersteller = (ta.Ersteller?.Title || '').toLowerCase();
                const erstellerNick = (ta.Ersteller?.Nickname || '').toLowerCase();
                const verantwortlicher = (ta.Verantwortlicher?.Title || '').toLowerCase();
                const verantwortlicherNick = (ta.Verantwortlicher?.Nickname || '').toLowerCase();
                const matchesSearch =
                    (ta.Title || '').toLowerCase().includes(s) ||
                    (ta.field_8 || '').toLowerCase().includes(s) ||
                    (ta.field_12 || '').toLowerCase().includes(s) ||
                    (ta.field_16 || '').toLowerCase().includes(s) ||
                    (ta.field_9 || '').toLowerCase().includes(s) ||
                    ersteller.includes(s) ||
                    erstellerNick.includes(s) ||
                    verantwortlicher.includes(s) ||
                    verantwortlicherNick.includes(s);
                if (!matchesSearch) return false;
            }
            // Status filter
            if (statusFilter === 'Offen' && ta.Status === 'abgeschlossen') return false;
            if (statusFilter !== 'Alle' && statusFilter !== 'Offen' && ta.Status !== statusFilter) return false;
            // Kategorie filter
            if (kategorieFilter !== 'Alle' && ta.field_16 !== kategorieFilter) return false;
            return true;
        });

        // Sort
        filtered.sort((a, b) => {
            const aVal = this.getSortValue(a, sortKey);
            const bVal = this.getSortValue(b, sortKey);
            const cmp = aVal.localeCompare(bVal, 'de');
            return sortDir === 'asc' ? cmp : -cmp;
        });

        return filtered;
    }

    private toggleSort(key: SortKey): void {
        this.setState(prev => ({
            sortKey: key,
            sortDir: prev.sortKey === key && prev.sortDir === 'asc' ? 'desc' : 'asc'
        }));
    }

    private renderSortIcon(key: SortKey): string {
        if (this.state.sortKey !== key) return ' ↕';
        return this.state.sortDir === 'asc' ? ' ↑' : ' ↓';
    }

    public render(): React.ReactElement<IAlleTasProps> {
        const filtered = this.getFilteredTas();
        const kategorien = this.getKategorien();

        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>Alle Technischen Anfragen</span>
                    </div>
                    <span className={styles.headerRight}>{filtered.length} Ergebnisse</span>
                </div>

                <div className={styles.content}>
                    {/* Filter Bar */}
                    <div className={styles.filterBar}>
                        <input
                            className={styles.searchInput}
                            placeholder="Suchen (TA-Nr., Kunde, Material, Kategorie, Ersteller)..."
                            value={this.state.search}
                            onChange={(e) => this.setState({ search: e.target.value })}
                        />
                        <select
                            className={styles.filterSelect}
                            value={this.state.statusFilter}
                            onChange={(e) => this.setState({ statusFilter: e.target.value })}
                        >
                            <option value="Offen">Offen (ohne abgeschlossen)</option>
                            <option value="Alle">Alle Status</option>
                            {STATUS_VALUES.map(s => (
                                <option key={s} value={s}>{s}</option>
                            ))}
                        </select>
                        <select
                            className={styles.filterSelect}
                            value={this.state.kategorieFilter}
                            onChange={(e) => this.setState({ kategorieFilter: e.target.value })}
                        >
                            <option value="Alle">Alle Kategorien</option>
                            {kategorien.map(k => (
                                <option key={k} value={k}>{k}</option>
                            ))}
                        </select>
                    </div>

                    {/* Table */}
                    {filtered.length === 0 ? (
                        <div className={styles.emptyState}>
                            <span className={styles.emptyIcon}></span>
                            <span>Keine TAs gefunden</span>
                        </div>
                    ) : (
                        <div style={{ overflowX: 'auto' }}>
                            <table className={styles.taTable}>
                                <thead>
                                    <tr>
                                        <th onClick={() => this.toggleSort('Title')}>TA-Nr.{this.renderSortIcon('Title')}</th>
                                        <th onClick={() => this.toggleSort('Ersteller')}>Ersteller{this.renderSortIcon('Ersteller')}</th>
                                        <th onClick={() => this.toggleSort('Kunde')}>Kunde{this.renderSortIcon('Kunde')}</th>
                                        <th onClick={() => this.toggleSort('Material')}>Material{this.renderSortIcon('Material')}</th>
                                        <th onClick={() => this.toggleSort('Kategorie')}>Kategorie{this.renderSortIcon('Kategorie')}</th>
                                        <th onClick={() => this.toggleSort('Verantwortlicher')}>Verantw.{this.renderSortIcon('Verantwortlicher')}</th>
                                        <th onClick={() => this.toggleSort('Termin')}>Termin{this.renderSortIcon('Termin')}</th>
                                        <th onClick={() => this.toggleSort('Status')}>Status{this.renderSortIcon('Status')}</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filtered.map(ta => (
                                        <tr key={ta.ID} onClick={() => this.props.onSelectTa(ta.ID)}>
                                            <td><b>{ta.Title}</b></td>
                                            <td title={ta.Ersteller?.Title}>{getKuerzel(ta.Ersteller)}</td>
                                            <td>{ta.field_8 || '–'}</td>
                                            <td>{ta.field_12 || '–'}</td>
                                            <td>{ta.field_16 || '–'}</td>
                                            <td title={ta.Verantwortlicher?.Title}>{getKuerzel(ta.Verantwortlicher)}</td>
                                            <td>{ta.field_6 ? new Date(ta.field_6).toLocaleDateString('de-DE') : '–'}</td>
                                            <td><StatusPill status={ta.Status || ''} /></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            </>
        );
    }
}
