import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem } from '../models/types';
import StatusPill from './StatusPill';

type SortKey = 'Title' | 'Ersteller' | 'Kunde' | 'Material' | 'Kategorie' | 'Endtermin' | 'Status';
type SortDir = 'asc' | 'desc';

const getInitials = (name?: string): string => {
    if (!name) return '–';
    const parts = name.trim().split(/\s+/);
    if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    return name.substring(0, 2).toUpperCase();
};

export interface ITerminPlanenProps {
    tas: ITaItem[];
    onSelectTa: (id: number) => void;
    onBack: () => void;
}

interface ITerminPlanenState {
    search: string;
    kategorieFilter: string;
    sortKey: SortKey;
    sortDir: SortDir;
}

export default class TerminPlanen extends React.Component<ITerminPlanenProps, ITerminPlanenState> {
    constructor(props: ITerminPlanenProps) {
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
        this.getPlanTas().forEach(ta => { if (ta.field_16) set.add(ta.field_16); });
        return Array.from(set).sort();
    }

    private getPlanTas(): ITaItem[] {
        return this.props.tas
            .filter(t => t.Status === 'Termin planen');
    }

    private getSortValue(ta: ITaItem, key: SortKey): string {
        switch (key) {
            case 'Title': return ta.Title || '';
            case 'Ersteller': return ta.Ersteller?.Title || '';
            case 'Kunde': return ta.field_8 || '';
            case 'Material': return ta.field_12 || '';
            case 'Kategorie': return ta.field_16 || '';
            case 'Endtermin': return ta.field_4 || '9999-12-31';
            case 'Status': return ta.Status || '';
            default: return '';
        }
    }

    private getFilteredTas(): ITaItem[] {
        const { search, kategorieFilter, sortKey, sortDir } = this.state;

        const filtered = this.getPlanTas().filter(ta => {
            if (search) {
                const s = search.toLowerCase();
                const ersteller = (ta.Ersteller?.Title || '').toLowerCase();
                const matchesSearch =
                    (ta.Title || '').toLowerCase().includes(s) ||
                    (ta.field_8 || '').toLowerCase().includes(s) ||
                    (ta.field_12 || '').toLowerCase().includes(s) ||
                    (ta.field_16 || '').toLowerCase().includes(s) ||
                    (ta.field_9 || '').toLowerCase().includes(s) ||
                    ersteller.includes(s);
                if (!matchesSearch) return false;
            }

            if (kategorieFilter !== 'Alle' && ta.field_16 !== kategorieFilter) return false;

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

    public render(): React.ReactElement<ITerminPlanenProps> {
        const filtered = this.getFilteredTas();
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
                            <table className={styles.taTable}>
                                <thead>
                                    <tr>
                                        <th onClick={() => this.toggleSort('Title')}>TA-Nr.{this.renderSortIcon('Title')}</th>
                                        <th onClick={() => this.toggleSort('Ersteller')}>Ersteller{this.renderSortIcon('Ersteller')}</th>
                                        <th onClick={() => this.toggleSort('Kunde')}>Kunde{this.renderSortIcon('Kunde')}</th>
                                        <th onClick={() => this.toggleSort('Material')}>Material{this.renderSortIcon('Material')}</th>
                                        <th onClick={() => this.toggleSort('Kategorie')}>Kategorie{this.renderSortIcon('Kategorie')}</th>
                                        <th onClick={() => this.toggleSort('Endtermin')}>Endtermin{this.renderSortIcon('Endtermin')}</th>
                                        <th onClick={() => this.toggleSort('Status')}>Status{this.renderSortIcon('Status')}</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {filtered.map(ta => (
                                        <tr key={ta.ID} onClick={() => this.props.onSelectTa(ta.ID)}>
                                            <td><b>{ta.Title}</b></td>
                                            <td title={ta.Ersteller?.Title}>{getInitials(ta.Ersteller?.Title)}</td>
                                            <td>{ta.field_8 || '–'}</td>
                                            <td>{ta.field_12 || '–'}</td>
                                            <td>{ta.field_16 || '–'}</td>
                                            <td>{this.formatDate(ta.field_4)}</td>
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
