import { StatusConfigService } from './StatusConfigService';
import { DEFAULT_STATUS_DEFINITIONS } from '../models/config';

describe('StatusConfigService', () => {
  afterEach(() => {
    StatusConfigService.reset();
  });

  it('defaults to the built-in status definitions', () => {
    expect(StatusConfigService.getAll()).toEqual(DEFAULT_STATUS_DEFINITIONS);
  });

  it('returns the label and color for a known default status', () => {
    expect(StatusConfigService.getLabel('überfällig')).toBe('überfällig');
    expect(StatusConfigService.getColor('überfällig')).toBe('#EF4444');
  });

  it('falls back to the key itself when a status has no configured label', () => {
    expect(StatusConfigService.getLabel('unknown-status')).toBe('unknown-status');
  });

  it('falls back to a neutral color when a status has no configured color', () => {
    expect(StatusConfigService.getColor('unknown-status')).toBe('#8a8f98');
  });

  it('uses tenant-configured definitions after init', () => {
    StatusConfigService.init([
      { key: 'open', label: 'Offen', color: '#123456' },
      { key: 'closed', label: 'Geschlossen', color: '#654321' }
    ]);

    expect(StatusConfigService.getLabel('open')).toBe('Offen');
    expect(StatusConfigService.getColor('closed')).toBe('#654321');
  });

  it('ignores an empty definitions array and keeps the default set', () => {
    StatusConfigService.init([]);
    expect(StatusConfigService.getAll()).toEqual(DEFAULT_STATUS_DEFINITIONS);
  });

  it('reset restores the default definitions', () => {
    StatusConfigService.init([{ key: 'open', label: 'Offen', color: '#123456' }]);
    StatusConfigService.reset();
    expect(StatusConfigService.getAll()).toEqual(DEFAULT_STATUS_DEFINITIONS);
  });
});
