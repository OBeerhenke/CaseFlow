import { LabelService } from './LabelService';
import { CONFIG_KEYS } from '../models/config';

describe('LabelService', () => {
  describe('getEntityLabelSingular / getEntityLabelPlural', () => {
    it('defaults to "Anfrage" / "Anfragen" when not configured', () => {
      expect(LabelService.getEntityLabelSingular({})).toBe('Anfrage');
      expect(LabelService.getEntityLabelPlural({})).toBe('Anfragen');
    });

    it('uses the customer-configured term when present', () => {
      const config = {
        [CONFIG_KEYS.ENTITY_LABEL_SINGULAR]: 'Ticket',
        [CONFIG_KEYS.ENTITY_LABEL_PLURAL]: 'Tickets'
      };
      expect(LabelService.getEntityLabelSingular(config)).toBe('Ticket');
      expect(LabelService.getEntityLabelPlural(config)).toBe('Tickets');
    });

    it('falls back to the default when the configured value is an empty string', () => {
      const config = { [CONFIG_KEYS.ENTITY_LABEL_SINGULAR]: '' };
      expect(LabelService.getEntityLabelSingular(config)).toBe('Anfrage');
    });
  });

  describe('getCreateActionLabel', () => {
    it('builds the default create-action label', () => {
      expect(LabelService.getCreateActionLabel({})).toBe('Neue Anfrage anlegen');
    });

    it('builds a customer-configured create-action label', () => {
      expect(LabelService.getCreateActionLabel({ [CONFIG_KEYS.ENTITY_LABEL_SINGULAR]: 'Case' }))
        .toBe('Neue Case anlegen');
    });
  });

  describe('getCreatedToastSuffix', () => {
    it('builds the default success-toast suffix', () => {
      expect(LabelService.getCreatedToastSuffix({})).toBe('Anfrage erfolgreich angelegt!');
    });
  });

  describe('getListViewTitle', () => {
    it('builds the default list view title using the plural form', () => {
      expect(LabelService.getListViewTitle({})).toBe('Alle Anfragen');
    });

    it('builds a customer-configured list view title', () => {
      expect(LabelService.getListViewTitle({ [CONFIG_KEYS.ENTITY_LABEL_PLURAL]: 'Vorgänge' }))
        .toBe('Alle Vorgänge');
    });
  });
});
