import React from 'react';

export const DEFAULT_ACCENT = '#3B82F6';
export const DEFAULT_ACCENT_RGB = '59, 130, 246';
export const DEFAULT_BG = '#f8fafc';

/**
 * Pure service that derives CSS custom property values from ConfigService
 * key-value data. No React state, no SharePoint dependency.
 */
export class ThemeService {
  /**
   * Build a React.CSSProperties object with `--cf-*` custom properties to
   * apply on the root `<div className={styles.caseFlowApp}>`.
   *
   * Usage from a component:
   *   <div className={styles.caseFlowApp} style={ThemeService.buildStyle(config)}>
   */
  public static buildStyle(config: Record<string, string>): React.CSSProperties {
    const accent = config.PrimaryColor || DEFAULT_ACCENT;
    const accentRgb = config.AccentColorRgb || DEFAULT_ACCENT_RGB;
    const bg = config.BgColor || DEFAULT_BG;

    return {
      '--cf-accent': accent,
      '--cf-accent-rgb': accentRgb,
      '--cf-bg': bg,
      '--cf-bg-card': 'rgba(255, 255, 255, 0.9)',
      '--cf-glass-bg': 'rgba(255, 255, 255, 0.85)',
      '--cf-glass-border': 'rgba(0, 0, 0, 0.08)',
      '--cf-text-main': 'rgb(15, 23, 42)',
      '--cf-text-muted': 'rgb(100, 116, 139)',
      backgroundColor: bg
    } as React.CSSProperties;
  }

  public static getLogoUrl(config: Record<string, string>): string | undefined {
    return config.LogoUrl || undefined;
  }

  public static getPrimaryColor(config: Record<string, string>): string {
    return config.PrimaryColor || DEFAULT_ACCENT;
  }
}
