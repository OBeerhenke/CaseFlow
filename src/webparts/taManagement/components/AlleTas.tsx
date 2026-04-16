import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaItem, STATUS_VALUES } from '../models/types';
import StatusPill from './StatusPill';

const prioLabel = (p?: number): { text: string; color: string } => {
    if (p === 3) return { text: '▲ Hoch', color: '#EF4444' };
    if (p === 2) return { text: '● Mittel', color: '#F59E0B' };
    if (p === 1) return { text: '▼ Niedrig', color: '#6B7280' };
    return { text: '–', color: '#CBD5E1' };
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
    prioFilter: string;
}

export default class AlleTas extends React.Component<IAlleTasProps, IAlleTasState> {
    constructor(props: IAlleTasProps) {
        super(props);
        this.state = {
            search: '',
            statusFilter: props.initialFilter || 'Alle',
            prioFilter: 'Alle'
        };
    }

    public componentDidUpdate(prevProps: IAlleTasProps): void {
        if (prevProps.initialFilter !== this.props.initialFilter && this.props.initialFilter) {
            this.setState({ statusFilter: this.props.initialFilter });
        }
    }

    private getFilteredTas(): ITaItem[] {
        const { search, statusFilter, prioFilter } = this.state;

        return this.props.tas
            .filter(ta => {
                // Text search
                if (search) {
                    const s = search.toLowerCase();
                    const ersteller = (ta.Ersteller?.Title || '').toLowerCase();
                    const matchesSearch =
                        (ta.Title || '').toLowerCase().includes(s) ||
                        (ta.field_8 || '').toLowerCase().includes(s) ||
                        (ta.field_12 || '').toLowerCase().includes(s) ||
                        (ta.field_9 || '').toLowerCase().includes(s) ||
                        ersteller.includes(s);
                    if (!matchesSearch) return false;
                }
                // Status filter
                if (statusFilter !== 'Alle' && ta.Status !== statusFilter) return false;
                // Priority filter
                if (prioFilter !== 'Alle' && ta.field_13 !== parseInt(prioFilter)) return false;
                return true;
            })
            .sort((a, b) => {
                // Sort by priority first, then by modified desc
                const prioDiff = (a.field_13 || 99) - (b.field_13 || 99);
                if (prioDiff !== 0) return prioDiff;
                return new Date(b.Modified || '').getTime() - new Date(a.Modified || '').getTime();
            });
    }

    public render(): React.ReactElement<IAlleTasProps> {
        const filtered = this.getFilteredTas();

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
                            placeholder="Suchen (TA-Nr., Kunde, Material, Ersteller)..."
                            value={this.state.search}
                            onChange={(e) => this.setState({ search: e.target.value })}
                        />
                        <select
                            className={styles.filterSelect}
                            value={this.state.statusFilter}
                            onChange={(e) => this.setState({ statusFilter: e.target.value })}
                        >
                            <option value="Alle">Alle Status</option>
                            {STATUS_VALUES.map(s => (
                                <option key={s} value={s}>{s}</option>
                            ))}
                        </select>
                    </div>

                    {/* List */}
                    {filtered.length === 0 ? (
                        <div className={styles.emptyState}>
                            <span className={styles.emptyIcon}></span>
                            <span>Keine TAs gefunden</span>
                        </div>
                    ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                            <div className={`${styles.listRow} ${styles.listRowAlleTas}`} style={{ background: 'transparent', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 12, textTransform: 'uppercase', letterSpacing: '0.05em', cursor: 'default', paddingBottom: 8 }}>
                                <span className={styles.listRowBold}>TA-Nr.</span>
                                <span className={styles.listRowMuted}>Kunde</span>
                                <span className={styles.listRowMuted}>Material</span>
                                <span className={styles.listRowMuted}>Verantwortlich</span>
                                <span className={styles.listRowMuted}>Prio</span>
                                <span className={styles.listRowMuted}>Verschiebung</span>
                                <span className={styles.listRowMuted}>Termin</span>
                                <span style={{ textAlign: 'center' }}>Status</span>
                            </div>
                            {filtered.map(ta => (
                                <div
                                    key={ta.ID}
                                    className={`${styles.listRow} ${styles.listRowAlleTas}`}
                                    onClick={() => this.props.onSelectTa(ta.ID)}
                                >
                                    <span className={styles.listRowBold}>{ta.Title}</span>
                                    <span className={styles.listRowMuted}>{ta.field_8 || '–'}</span>
                                    <span className={styles.listRowMuted}>{ta.field_12 || '–'}</span>
                                    <span className={styles.listRowMuted} title={ta.Verantwortlicher?.Title}>{ta.Verantwortlicher?.Title || '–'}</span>
                                    <span style={{ fontSize: 11, fontWeight: 600, color: prioLabel(ta.field_13).color }}>{prioLabel(ta.field_13).text}</span>
                                    <span className={styles.listRowWarn}>{ta.field_21 || '–'}</span>
                                    <span className={styles.listRowMuted}>{ta.field_6 || '–'}</span>
                                    <div style={{ display: 'flex', justifyContent: 'center' }}>
                                        <StatusPill status={ta.Status || ''} />
                                    </div>
                                </div>
                            ))}
                        </div>
                    )}
                </div>
            </>
        );
    }
}
