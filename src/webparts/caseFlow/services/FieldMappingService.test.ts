import { FieldMappingService } from './FieldMappingService';
import { DEFAULT_CASE_FIELD_MAPPING } from '../models/config';

describe('FieldMappingService', () => {
  afterEach(() => {
    FieldMappingService.reset();
  });

  describe('resolve', () => {
    it('resolves a normalized key to its default (identity) SharePoint field name', () => {
      expect(FieldMappingService.resolve('field_8')).toBe('field_8');
      expect(FieldMappingService.resolve('Status')).toBe('Status');
    });

    it('resolves a normalized key to a tenant-configured override', () => {
      FieldMappingService.init({ field_8: 'Customer' });
      expect(FieldMappingService.resolve('field_8')).toBe('Customer');
    });

    it('falls back to the key itself when it is not part of the mapping', () => {
      expect(FieldMappingService.resolve('SomeUnmappedKey')).toBe('SomeUnmappedKey');
    });

    it('falls back to the default mapping for keys not present in the override', () => {
      FieldMappingService.init({ field_8: 'Customer' });
      expect(FieldMappingService.resolve('field_6')).toBe('field_6');
    });
  });

  describe('resolveMany', () => {
    it('resolves a list of normalized keys', () => {
      FieldMappingService.init({ field_8: 'Customer', field_6: 'PlannedDate' });
      expect(FieldMappingService.resolveMany(['field_8', 'field_6', 'Status']))
        .toEqual(['Customer', 'PlannedDate', 'Status']);
    });
  });

  describe('denormalize', () => {
    it('translates a raw SharePoint item (default mapping) into the normalized shape', () => {
      const raw = { ID: 1, Title: 'CF-2026-001', field_8: 'Acme GmbH', Status: 'Termin planen' };
      const result = FieldMappingService.denormalize<Record<string, unknown>>(raw);
      expect(result).toEqual({ ID: 1, Title: 'CF-2026-001', field_8: 'Acme GmbH', Status: 'Termin planen' });
    });

    it('translates a raw item using a tenant-specific field mapping back to normalized keys', () => {
      FieldMappingService.init({ field_8: 'Customer', field_6: 'PlannedDate' });
      const raw = { ID: 1, Title: 'CF-2026-001', Customer: 'Acme GmbH', PlannedDate: '2026-09-01', Status: 'läuft planmäßig' };
      const result = FieldMappingService.denormalize<Record<string, unknown>>(raw);
      expect(result.field_8).toBe('Acme GmbH');
      expect(result.field_6).toBe('2026-09-01');
      expect(result.Status).toBe('läuft planmäßig');
      // Raw tenant-specific keys should not leak into the normalized shape
      expect(result).not.toHaveProperty('Customer');
      expect(result).not.toHaveProperty('PlannedDate');
    });

    it('drops fields that are missing from the raw item entirely', () => {
      const raw = { ID: 1, Title: 'CF-2026-001' };
      const result = FieldMappingService.denormalize<Record<string, unknown>>(raw);
      expect(result).toEqual({ ID: 1, Title: 'CF-2026-001' });
      expect(result).not.toHaveProperty('field_8');
    });
  });

  describe('normalize', () => {
    it('translates a normalized update payload into default SharePoint field names', () => {
      const payload = { field_6: '2026-09-01', Status: 'läuft planmäßig' };
      expect(FieldMappingService.normalize(payload)).toEqual({
        field_6: '2026-09-01',
        Status: 'läuft planmäßig'
      });
    });

    it('translates a normalized update payload using a tenant-specific field mapping', () => {
      FieldMappingService.init({ field_6: 'PlannedDate', field_21: 'DelayReason' });
      const payload = { field_6: '2026-09-01', field_21: 'Kapazität' };
      expect(FieldMappingService.normalize(payload)).toEqual({
        PlannedDate: '2026-09-01',
        DelayReason: 'Kapazität'
      });
    });

    it('passes through keys that are not part of the field mapping unchanged', () => {
      const payload = { ErstellerId: 5, VerantwortlicherId: 12 };
      expect(FieldMappingService.normalize(payload)).toEqual({
        ErstellerId: 5,
        VerantwortlicherId: 12
      });
    });
  });

  describe('current', () => {
    it('returns a copy of the active mapping, defaulting to DEFAULT_CASE_FIELD_MAPPING', () => {
      expect(FieldMappingService.current).toEqual(DEFAULT_CASE_FIELD_MAPPING);
    });

    it('returned mapping is a copy — mutating it does not affect the service', () => {
      const snapshot = FieldMappingService.current;
      snapshot.field_8 = 'Mutated';
      expect(FieldMappingService.resolve('field_8')).toBe('field_8');
    });

    it('reflects overrides after init', () => {
      FieldMappingService.init({ field_8: 'Customer' });
      expect(FieldMappingService.current.field_8).toBe('Customer');
      expect(FieldMappingService.current.field_6).toBe('field_6');
    });
  });

  describe('reset', () => {
    it('restores the default mapping', () => {
      FieldMappingService.init({ field_8: 'Customer' });
      FieldMappingService.reset();
      expect(FieldMappingService.resolve('field_8')).toBe('field_8');
    });
  });
});
