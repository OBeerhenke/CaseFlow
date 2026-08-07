import * as React from 'react';
import styles from './App.module.scss';
import { ICaseItem } from '../models/types';

export interface IAnalyticsProps {
    cases: ICaseItem[];
    onBack: () => void;
    onSelectCase: (id: number) => void;
}

interface IAnalyticsState {
    zeitraum: 'alle' | 'jahr' | 'quartal';
    jahr: number;
}

/** Workday diff between two dates (Mon–Fri only) */
function workdayDiff(from: Date, to: Date): number {
    let count = 0;
    const step = from <= to ? 1 : -1;
    const d = new Date(from);
    d.setHours(0, 0, 0, 0);
    const target = new Date(to);
    target.setHours(0, 0, 0, 0);

    while ((step > 0 && d < target) || (step < 0 && d > target)) {
        d.setDate(d.getDate() + step);
        const dow = d.getDay();
        if (dow !== 0 && dow !== 6) count += step;
    }
    return count;
}

function parseDate(s?: string): Date | null {
    if (!s) return null;
    if (s.includes('-')) {
        const d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }
    const pts = s.split('.');
    if (pts.length === 3) {
        const d = new Date(parseInt(pts[2]), parseInt(pts[1]) - 1, parseInt(pts[0]));
        return isNaN(d.getTime()) ? null : d;
    }
    return null;
}

interface ITermintreueRow {
    caseItem: ICaseItem;
    istVsWunsch: number | null;      // IST-Erledigung vs. Wunschtermin
    wunschVsPlan: number | null;     // Wunschtermin vs. geplanter Termin
    istVsOriginal: number | null;    // IST vs. ursprünglich geplanter Termin
    istVsVerschoben: number | null;  // IST vs. verschobener (aktueller) Termin
}

function computeRow(caseItem: ICaseItem): ITermintreueRow {
    // Fallback: wenn kein Erledigungsdatum gesetzt ist (Alt-TAs), nutze Modified
    const erledigung = parseDate(caseItem.Erledigungsdatum) || (caseItem.Status === 'abgeschlossen' ? parseDate(caseItem.Modified) : null);
    const wunsch = parseDate(caseItem.field_5);
    const geplant = parseDate(caseItem.field_6);
    const original = parseDate(caseItem.field_22);

    return {
        caseItem,
        istVsWunsch: erledigung && wunsch ? workdayDiff(erledigung, wunsch) : null,
        wunschVsPlan: wunsch && geplant ? workdayDiff(geplant, wunsch) : null,
        istVsOriginal: erledigung && original ? workdayDiff(erledigung, original) : null,
        istVsVerschoben: erledigung && geplant ? workdayDiff(erledigung, geplant) : null,
    };
}

function avg(values: number[]): number {
    if (values.length === 0) return 0;
    return values.reduce((a, b) => a + b, 0) / values.length;
}

function pct(count: number, total: number): string {
    if (total === 0) return '0';
    return ((count / total) * 100).toFixed(0);
}

/** Color-coded diff cell: positive = green (ahead), negative = red (late) */
const DiffCell: React.FC<{ value: number | null }> = ({ value }) => {
    if (value === null) return <span className={styles.statsColNum} style={{ color: '#94a3b8' }}>–</span>;

    const color = value >= 0 ? '#10B981' : '#EF4444';
    const prefix = value > 0 ? '+' : '';

    return (
        <span className={styles.statsColNum} style={{ color, fontWeight: 600 }}>
            {prefix}{value} WT
        </span>
    );
};

export default class Analytics extends React.Component<IAnalyticsProps, IAnalyticsState> {
    constructor(props: IAnalyticsProps) {
        super(props);
        this.state = {
            zeitraum: 'alle',
            jahr: new Date().getFullYear()
        };
    }

    private getFilteredCases(): ICaseItem[] {
        const { zeitraum, jahr } = this.state;
        const completed = this.props.cases.filter(t => t.Status === 'abgeschlossen');

        if (zeitraum === 'alle') return completed;

        return completed.filter(t => {
            const d = parseDate(t.Erledigungsdatum) || parseDate(t.field_4);
            if (!d) return false;
            if (d.getFullYear() !== jahr) return false;
            if (zeitraum === 'quartal') {
                const now = new Date();
                const q = Math.floor(now.getMonth() / 3);
                const tq = Math.floor(d.getMonth() / 3);
                return tq === q;
            }
            return true;
        });
    }

    private getStatusDistribution(): { label: string; count: number; color: string; pct: string }[] {
        const all = this.props.cases;
        const total = all.length;
        return [
            { label: 'Überfällig', count: all.filter(t => t.Status === 'überfällig').length, color: '#EF4444', pct: '' },
            { label: 'Prüfen', count: all.filter(t => t.Status === 'prüfen').length, color: '#F59E0B', pct: '' },
            { label: 'Planmäßig', count: all.filter(t => t.Status === 'läuft planmäßig').length, color: '#10B981', pct: '' },
            { label: 'Termin planen', count: all.filter(t => t.Status === 'Termin planen').length, color: '#3B82F6', pct: '' },
            { label: 'Abgeschlossen', count: all.filter(t => t.Status === 'abgeschlossen').length, color: '#6B7280', pct: '' },
        ].map(s => ({ ...s, pct: pct(s.count, total) }));
    }

    private getVerschiebungsStats(): { grund: string; count: number }[] {
        const gruende = new Map<string, number>();
        this.props.cases.forEach(t => {
            if (t.field_21) {
                gruende.set(t.field_21, (gruende.get(t.field_21) || 0) + 1);
            }
        });
        return Array.from(gruende.entries())
            .map(([grund, count]) => ({ grund, count }))
            .sort((a, b) => b.count - a.count);
    }

    public render(): React.ReactElement<IAnalyticsProps> {
        const filtered = this.getFilteredCases();
        const rows = filtered.map(computeRow);
        const statusDist = this.getStatusDistribution();
        const verschiebungen = this.getVerschiebungsStats();

        // Aggregated metrics
        const istVsWunschVals = rows.map(r => r.istVsWunsch).filter(v => v !== null) as number[];
        const istVsOriginalVals = rows.map(r => r.istVsOriginal).filter(v => v !== null) as number[];
        const istVsVerschobenVals = rows.map(r => r.istVsVerschoben).filter(v => v !== null) as number[];

        const termintreuePct = istVsVerschobenVals.length > 0
            ? pct(istVsVerschobenVals.filter(v => v >= 0).length, istVsVerschobenVals.length)
            : '–';

        const avgAbweichungWunsch = istVsWunschVals.length > 0 ? avg(istVsWunschVals).toFixed(1) : '–';
        const avgAbweichungPlan = istVsVerschobenVals.length > 0 ? avg(istVsVerschobenVals).toFixed(1) : '–';

        const totalCases = this.props.cases.length;
        const completedCount = filtered.length;
        const verschobenCount = this.props.cases.filter(t => t.field_22).length;

        return (
            <>
                <div className={styles.header}>
                    <div className={styles.headerWithBack}>
                        <button className={styles.backButton} onClick={this.props.onBack}>← Zurück</button>
                        <span className={styles.headerTitle}>Analytics & Termintreue</span>
                    </div>
                </div>

                <div className={styles.content}>
                    {/* Filter Bar */}
                    <div className={styles.filterBar}>
                        <select
                            className={styles.filterSelect}
                            value={this.state.zeitraum}
                            onChange={(e) => this.setState({ zeitraum: e.target.value as any })}
                        >
                            <option value="alle">Alle Zeiträume</option>
                            <option value="jahr">Aktuelles Jahr</option>
                            <option value="quartal">Aktuelles Quartal</option>
                        </select>
                    </div>

                    {/* ══════ Overview KPIs ══════ */}
                    <div className={styles.statsGrid}>
                        <div className={styles.statCard}>
                            <span className={styles.statValue} style={{ color: '#10B981' }}>
                                {termintreuePct === '–' ? '–' : `${termintreuePct}%`}
                            </span>
                            <span className={styles.statLabel}>Termintreue</span>
                            <span className={styles.statSub}>Pünktlich / gepl. Termin</span>
                        </div>
                        <div className={styles.statCard}>
                            <span className={styles.statValue} style={{ color: '#3B82F6' }}>{completedCount}</span>
                            <span className={styles.statLabel}>Abgeschlossen</span>
                            <span className={styles.statSub}>von {totalCases} gesamt</span>
                        </div>
                        <div className={styles.statCard}>
                            <span className={styles.statValue} style={{ color: '#F59E0B' }}>{verschobenCount}</span>
                            <span className={styles.statLabel}>Verschoben</span>
                            <span className={styles.statSub}>mind. 1× umgeplant</span>
                        </div>
                        <div className={styles.statCard}>
                            <span className={styles.statValue} style={{
                                color: avgAbweichungPlan === '–' ? '#94a3b8' :
                                    parseFloat(avgAbweichungPlan.toString()) >= 0 ? '#10B981' : '#EF4444'
                            }}>
                                {avgAbweichungPlan === '–' ? '–' : `${avgAbweichungPlan} WT`}
                            </span>
                            <span className={styles.statLabel}>Ø Abweichung</span>
                            <span className={styles.statSub}>IST vs. Plan (Werktage)</span>
                        </div>
                    </div>

                    {/* ══════ Avg Metrics Detail ══════ */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Durchschnittliche Abweichungen (Werktage)</h4>
                        <div className={styles.metricsRow}>
                            <div className={styles.metricPill}>
                                <span className={styles.metricLabel}>IST vs. Wunsch</span>
                                <span className={styles.metricValue} style={{
                                    color: parseFloat(avgAbweichungWunsch.toString()) >= 0 ? '#10B981' : '#EF4444'
                                }}>
                                    {avgAbweichungWunsch} WT
                                </span>
                            </div>
                            <div className={styles.metricPill}>
                                <span className={styles.metricLabel}>IST vs. Plan</span>
                                <span className={styles.metricValue} style={{
                                    color: parseFloat(avgAbweichungPlan.toString()) >= 0 ? '#10B981' : '#EF4444'
                                }}>
                                    {avgAbweichungPlan} WT
                                </span>
                            </div>
                            <div className={styles.metricPill}>
                                <span className={styles.metricLabel}>IST vs. Original</span>
                                <span className={styles.metricValue} style={{
                                    color: istVsOriginalVals.length > 0 && avg(istVsOriginalVals) >= 0 ? '#10B981' : '#EF4444'
                                }}>
                                    {istVsOriginalVals.length > 0 ? avg(istVsOriginalVals).toFixed(1) : '–'} WT
                                </span>
                            </div>
                        </div>
                    </div>

                    {/* ══════ Status Distribution ══════ */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>Statusverteilung</h4>
                        <div className={styles.barChartContainer}>
                            {/* Stacked bar */}
                            <div className={styles.stackedBar}>
                                {statusDist.filter(s => s.count > 0).map(s => (
                                    <div
                                        key={s.label}
                                        className={styles.stackedSegment}
                                        style={{
                                            width: `${pct(s.count, totalCases)}%`,
                                            backgroundColor: s.color
                                        }}
                                        title={`${s.label}: ${s.count} (${s.pct}%)`}
                                    />
                                ))}
                            </div>
                            {/* Legend */}
                            <div className={styles.chartLegend}>
                                {statusDist.map(s => (
                                    <div key={s.label} className={styles.legendItem}>
                                        <span className={styles.legendDot} style={{ backgroundColor: s.color }} />
                                        <span className={styles.legendLabel}>{s.label}</span>
                                        <span className={styles.legendCount}>{s.count} ({s.pct}%)</span>
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>

                    {/* ══════ Verschiebungsgründe ══════ */}
                    {verschiebungen.length > 0 && (
                        <div className={styles.formGroup}>
                            <h4 className={styles.formGroupTitle}>Häufigste Verschiebungsgründe</h4>
                            <div className={styles.reasonsList}>
                                {verschiebungen.slice(0, 8).map((v, i) => {
                                    const maxCount = verschiebungen[0].count;
                                    return (
                                        <div key={i} className={styles.reasonRow}>
                                            <span className={styles.reasonLabel}>{v.grund}</span>
                                            <div className={styles.reasonBarTrack}>
                                                <div
                                                    className={styles.reasonBarFill}
                                                    style={{ width: `${(v.count / maxCount) * 100}%` }}
                                                />
                                            </div>
                                            <span className={styles.reasonCount}>{v.count}</span>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>
                    )}

                    {/* ══════ Detail Table ══════ */}
                    <div className={styles.formGroup}>
                        <h4 className={styles.formGroupTitle}>
                            Termintreue-Details ({rows.length} abgeschlossene TAs)
                        </h4>

                        {rows.length === 0 ? (
                            <div className={styles.emptyState}>
                                <span className={styles.emptyIcon} />
                                <span>Keine abgeschlossenen TAs im gewählten Zeitraum</span>
                            </div>
                        ) : (
                            <div className={styles.statsTable}>
                                <div className={styles.statsTableHead}>
                                    <span className={styles.statsColCase}>TA-Nr.</span>
                                    <span className={styles.statsColKunde}>Kunde</span>
                                    <span className={styles.statsColNum}>IST↔Wunsch</span>
                                    <span className={styles.statsColNum}>Wunsch↔Plan</span>
                                    <span className={styles.statsColNum}>IST↔Original</span>
                                    <span className={styles.statsColNum}>IST↔Plan</span>
                                </div>
                                {rows.map(r => (
                                    <div
                                        key={r.caseItem.ID}
                                        className={styles.statsTableRow}
                                        onClick={() => this.props.onSelectCase(r.caseItem.ID)}
                                    >
                                        <span className={styles.statsColCase}>{r.caseItem.Title}</span>
                                        <span className={styles.statsColKunde}>{r.caseItem.field_8 || '–'}</span>
                                        <DiffCell value={r.istVsWunsch} />
                                        <DiffCell value={r.wunschVsPlan} />
                                        <DiffCell value={r.istVsOriginal} />
                                        <DiffCell value={r.istVsVerschoben} />
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            </>
        );
    }
}
