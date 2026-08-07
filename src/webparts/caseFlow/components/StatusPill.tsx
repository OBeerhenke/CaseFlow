import * as React from 'react';
import styles from './App.module.scss';
import { StatusConfigService } from '../services/StatusConfigService';

export interface IStatusPillProps {
    status: string;
}

/** Convert a "#rrggbb" hex color into an "rgba(r, g, b, alpha)" string. */
function hexToRgba(hex: string, alpha: number): string {
    const clean = hex.replace('#', '');
    if (clean.length !== 6) return hex;
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    if ([r, g, b].some(isNaN)) return hex;
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

/**
 * Renders a status as a colored pill. Label and color come from
 * `StatusConfigService`, which is tenant-configurable via the Property
 * Pane — this component itself has no hardcoded status/color mapping.
 */
const StatusPill: React.FC<IStatusPillProps> = ({ status }) => {
    const color = StatusConfigService.getColor(status);
    const label = StatusConfigService.getLabel(status);

    return (
        <span
            className={styles.statusPill}
            style={{ color, background: hexToRgba(color, 0.15) }}
        >
            {label}
        </span>
    );
};

export default StatusPill;
