import * as React from 'react';
import styles from './App.module.scss';
import { ICaseItem, STATUS_VALUES } from '../models/types';
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

export interface ICaseListProps {
    cases: ICaseItem[];
    initialFilter: string;
    onSelectCase: (id: number) => void;
    onBack: () => void;
}

interface ICaseListState {
    search: string;
    statusFilter: string;
    kategorieFilter: string;
    sortKey: SortKey;
    sortDir: SortDir;
}

export default class CaseList extends React.Component<ICaseListProps, ICaseListState> {
    constructor(props: ICaseListProps) {
        super(props);
        this.state = {
            search: '',
            statusFilter: props.initialFilter || 'Offen',
            kategorieFilter: 'Alle',
            sortKey: 'Termin',
            sortDir: 'asc'
        };
    }

    public componentDidUpdate(prevProps: ICaseListProps): void {
        if (prevProps.initialFilter !== this.props.initialFilter && this.props.initialFilter) {
            this.setState({ statusFilter: this.props.initialFilter });
        }
    }

    private getKategorien(): string[] {
        const set = new Set<string>();
        this.props.cases.forEach(caseItem => { if (caseItem.field_16) set.add(caseItem.field_16); });
        return Array.from(set).sort();
    }

    private getSortValue(caseItem: ICaseItem, key: SortKey): string {
        switch (key) {
            case 'Title': return caseItem.Title || '';
            case 'Ersteller': return caseItem.Ersteller?.Title || '';
            case 'Kunde': return caseItem.field_8 || '';
            case 'Material': return caseItem.field_12 || '';
            case 'Kategorie': return caseItem.field_16 || '';
            case 'Verantwortlicher': return caseItem.Verantwortlicher?.Title || '';
            case 'Termin': return caseItem.field_6 || '9999-12-31';
            case 'Status': return caseItem.Status || '';
            default: return '';
        }
    }

    private getFilteredCases(): ICaseItem[] {
        const { search, statusFilter, kategorieFilter, sortKey, sortDir } = this.state;

        const filtered = this.props.cases.filter(caseItem => {
            // Text search
            if (search) {
                const s = search.toLowerCase();
                const ersteller = (caseItem.Ersteller?.Title || '').toLowerCase();
                const erstellerNick = (caseItem.Ersteller?.Nickname || '').toLowerCase();
                const verantwortlicher = (caseItem.Verantwortlicher?.Title || '').toLowerCase();
                const verantwortlicherNick = (caseItem.Verantwortlicher?.Nickname || '').toLowerCase();
                const matchesSearch =
                    (caseItem.Title || '').toLowerCase().includes(s) ||
                    (caseItem.field_8 || '').toLowerCase().includes(s) ||
                    (caseItem.field_12 || '').toLowerCase().includes(s) ||
                    (caseItem.field_16 || '').toLowerCase().includes(s) ||
                    (caseItem.field_9 || '').toLowerCase().includes(s) ||
                    ersteller.includes(s) ||
                    erstellerNick.includes(s) ||
                    verantwortlicher.includes(s) ||
                    verantwortlicherNick.includes(s);
                if (!matchesSearch) return false;
            }
            // Status filter
            if (statusFilter === 'Offen' && caseItem.Status === 'abgeschlossen') return false;
            if (statusFilter !== 'Alle' && statusFilter !== 'Offen' && caseItem.Status !== statusFilter) return false;
            // Kategorie filter
            if (kategorieFilter !== 'Alle' && caseItem.field_16 !== kategorieFilter) return false;
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

    public render(): React.ReactElement<ICaseListProps> {
        const filtered = this.getFilteredCases();
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
                            <table className={styles.caseTable}>
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
                                    {filtered.map(caseItem => (
                                        <tr key={caseItem.ID} onClick={() => this.props.onSelectCase(caseItem.ID)}>
                                            <td><b>{caseItem.Title}</b></td>
                                            <td title={caseItem.Ersteller?.Title}>{getKuerzel(caseItem.Ersteller)}</td>
                                            <td>{caseItem.field_8 || '–'}</td>
                                            <td>{caseItem.field_12 || '–'}</td>
                                            <td>{caseItem.field_16 || '–'}</td>
                                            <td title={caseItem.Verantwortlicher?.Title}>{getKuerzel(caseItem.Verantwortlicher)}</td>
                                            <td>{caseItem.field_6 ? new Date(caseItem.field_6).toLocaleDateString('de-DE') : '–'}</td>
                                            <td><StatusPill status={caseItem.Status || ''} /></td>
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
