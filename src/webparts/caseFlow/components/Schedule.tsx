import * as React from 'react';
import styles from './App.module.scss';
import { ICaseItem } from '../models/types';
import StatusPill from './StatusPill';

type SortKey = 'Title' | 'Ersteller' | 'Kunde' | 'Material' | 'Kategorie' | 'Endtermin' | 'Status';
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

export interface IScheduleProps {
    cases: ICaseItem[];
    onSelectCase: (id: number) => void;
    onBack: () => void;
}

interface IScheduleState {
    search: string;
    kategorieFilter: string;
    sortKey: SortKey;
    sortDir: SortDir;
}

export default class Schedule extends React.Component<IScheduleProps, IScheduleState> {
    constructor(props: IScheduleProps) {
        super(props);
        this.state = {
            search: '',
            kategorieFilter: 'Alle',
            sortKey: 'Endtermin',
            sortDir: 'asc'
        };
    }

    private getKategorien(): string[] {
        const set = new Set<string>();
        this.getPlanCases().forEach(caseItem => { if (caseItem.field_16) set.add(caseItem.field_16); });
        return Array.from(set).sort();
    }

    private getPlanCases(): ICaseItem[] {
        return this.props.cases
            .filter(t => t.Status === 'Termin planen');
    }

    private getSortValue(caseItem: ICaseItem, key: SortKey): string {
        switch (key) {
            case 'Title': return caseItem.Title || '';
            case 'Ersteller': return caseItem.Ersteller?.Title || '';
            case 'Kunde': return caseItem.field_8 || '';
            case 'Material': return caseItem.field_12 || '';
            case 'Kategorie': return caseItem.field_16 || '';
            case 'Endtermin': return caseItem.field_4 || '9999-12-31';
            case 'Status': return caseItem.Status || '';
            default: return '';
        }
    }

    private getFilteredCases(): ICaseItem[] {
        const { search, kategorieFilter, sortKey, sortDir } = this.state;

        const filtered = this.getPlanCases().filter(caseItem => {
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

            if (kategorieFilter !== 'Alle' && caseItem.field_16 !== kategorieFilter) return false;

            return true;
        });

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

    private formatDate(dateValue?: string): string {
        if (!dateValue) return '–';
        const d = new Date(dateValue);
        if (!isNaN(d.getTime())) return d.toLocaleDateString('de-DE');

        const parts = dateValue.split('.');
        if (parts.length === 3) return dateValue;
        return '–';
    }

    public render(): React.ReactElement<IScheduleProps> {
        const filtered = this.getFilteredCases();
        const kategorien = this.getKategorien();

        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>TA Planen</span>
                    </div>
                    <span className={styles.headerRight}>{filtered.length} Ergebnisse</span>
                </div>

                <div className={styles.content}>
                    <div className={styles.filterBar}>
                        <input
                            className={styles.searchInput}
                            placeholder="Suchen (TA-Nr., Kunde, Material, Kategorie, Ersteller)..."
                            value={this.state.search}
                            onChange={(e) => this.setState({ search: e.target.value })}
                        />
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

                    {filtered.length === 0 ? (
                        <div className={styles.emptyState}>
                            <span className={styles.emptyIcon}></span>
                            <span>Alle Termine sind geplant!</span>
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
                                        <th onClick={() => this.toggleSort('Endtermin')}>Erstelltermin{this.renderSortIcon('Endtermin')}</th>
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
                                            <td>{this.formatDate(caseItem.field_4)}</td>
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
