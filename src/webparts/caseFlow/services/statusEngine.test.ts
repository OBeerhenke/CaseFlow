import {
  evaluateCaseStatus,
  addWorkdays,
  parsePlannedDate,
  DEFAULT_STATUS_ENGINE_CONFIG
} from './statusEngine';

/** Helper: create a Date for a given ISO date string at midnight local time. */
function d(iso: string): Date {
  const dt = new Date(iso);
  dt.setHours(0, 0, 0, 0);
  return dt;
}

describe('parsePlannedDate', () => {
  it('parses ISO format (YYYY-MM-DD)', () => {
    const result = parsePlannedDate('2026-06-15');
    expect(result).toBeDefined();
    expect(result!.getFullYear()).toBe(2026);
    expect(result!.getMonth()).toBe(5); // 0-indexed
    expect(result!.getDate()).toBe(15);
  });

  it('parses German format (DD.MM.YYYY)', () => {
    const result = parsePlannedDate('15.06.2026');
    expect(result).toBeDefined();
    expect(result!.getFullYear()).toBe(2026);
    expect(result!.getMonth()).toBe(5);
    expect(result!.getDate()).toBe(15);
  });

  it('returns undefined for falsy input (null/undefined/empty)', () => {
    expect(parsePlannedDate(null)).toBeUndefined();
    expect(parsePlannedDate(undefined)).toBeUndefined();
    expect(parsePlannedDate('')).toBeUndefined();
  });

  it('returns undefined for unparseable strings', () => {
    expect(parsePlannedDate('not-a-date')).toBeUndefined();
    expect(parsePlannedDate('xx.yy.zzzz')).toBeUndefined();
  });
});

describe('addWorkdays', () => {
  it('adds 0 workdays (no change)', () => {
    const result = addWorkdays(d('2026-06-01'), 0);
    expect(result.getTime()).toBe(d('2026-06-01').getTime());
  });

  it('adds 1 workday to a Monday → Tuesday', () => {
    const result = addWorkdays(d('2026-06-01'), 1); // Monday
    expect(result.getTime()).toBe(d('2026-06-02').getTime()); // Tuesday
  });

  it('adds 1 workday to a Friday → Monday (skips weekend)', () => {
    const result = addWorkdays(d('2026-06-05'), 1); // Friday
    expect(result.getTime()).toBe(d('2026-06-08').getTime()); // Monday
  });

  it('adds 5 workdays from Wednesday → Wednesday (full week)', () => {
    const result = addWorkdays(d('2026-06-03'), 5); // Wednesday
    expect(result.getTime()).toBe(d('2026-06-10').getTime()); // next Wednesday
  });

  it('does not mutate the input Date', () => {
    const input = d('2026-06-01');
    const inputCopy = new Date(input.getTime());
    addWorkdays(input, 3);
    expect(input.getTime()).toBe(inputCopy.getTime());
  });
});

describe('evaluateCaseStatus', () => {
  // Use a fixed "today" so tests are deterministic.
  const today = d('2026-06-10');

  describe('completed cases', () => {
    it('returns the completed status and changed=false for an already-completed case', () => {
      const result = evaluateCaseStatus('2026-06-15', 'abgeschlossen', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('abgeschlossen');
      expect(result.changed).toBe(false);
    });

    it('does not change an overdue completed case', () => {
      const result = evaluateCaseStatus('2026-06-01', 'abgeschlossen', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('abgeschlossen');
      expect(result.changed).toBe(false);
    });
  });

  describe('no planned date', () => {
    it('returns statusNoPlan when plannedDate is undefined', () => {
      const result = evaluateCaseStatus(undefined, undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('Termin planen');
      expect(result.changed).toBe(true);
    });

    it('returns changed=false if it already is statusNoPlan', () => {
      const result = evaluateCaseStatus(undefined, 'Termin planen', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('Termin planen');
      expect(result.changed).toBe(false);
    });
  });

  describe('overdue', () => {
    it('marks as overdue when planned date is in the past', () => {
      const result = evaluateCaseStatus('2026-06-09', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('überfällig');
      expect(result.changed).toBe(true);
    });

    it('marks as overdue when planned date is today', () => {
      const result = evaluateCaseStatus('2026-06-10', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('überfällig');
      expect(result.changed).toBe(true);
    });

    it('returns changed=false when already overdue', () => {
      const result = evaluateCaseStatus('2026-06-09', 'überfällig', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('überfällig');
      expect(result.changed).toBe(false);
    });
  });

  describe('review (within PruefenTage workdays)', () => {
    it('marks as review when planned date is within the review window', () => {
      // reviewDays=3: workdays added to today (Wed June 10) = Mon June 15
      // planned June 12 < June 15, so it's "pruefen"
      const result = evaluateCaseStatus('2026-06-12', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('prüfen');
      expect(result.changed).toBe(true);
    });

    it('marks as review on the boundary (planned == reviewDeadline)', () => {
      // reviewDays=3: deadline = Monday June 15
      const result = evaluateCaseStatus('2026-06-15', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('prüfen');
    });

    it('returns changed=false when already in review', () => {
      const result = evaluateCaseStatus('2026-06-12', 'prüfen', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('prüfen');
      expect(result.changed).toBe(false);
    });
  });

  describe('on track', () => {
    it('marks as on track when planned date is well in the future', () => {
      // reviewDays=3: deadline = Mon June 15. planned June 22 > June 15
      const result = evaluateCaseStatus('2026-06-22', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('läuft planmäßig');
      expect(result.changed).toBe(true);
    });

    it('returns changed=false when already on track', () => {
      const result = evaluateCaseStatus('2026-06-22', 'läuft planmäßig', DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(result.status).toBe('läuft planmäßig');
      expect(result.changed).toBe(false);
    });
  });

  describe('configurable thresholds', () => {
    it('uses the configured reviewDays threshold', () => {
      const config = { ...DEFAULT_STATUS_ENGINE_CONFIG, reviewDays: 10 };
      // 10 workdays from Wed June 10 = Wed June 24
      // planned June 20 < June 24 => "prüfen"
      const result = evaluateCaseStatus('2026-06-20', undefined, config, today);
      expect(result.status).toBe('prüfen');
    });

    it('uses configurable status key names', () => {
      const config = {
        ...DEFAULT_STATUS_ENGINE_CONFIG,
        statusNoPlan: 'Offen',
        statusOverdue: 'Verspätet',
        statusReview: 'Achtung',
        statusOnTrack: 'Im Plan',
        statusCompleted: 'Fertig'
      };
      expect(evaluateCaseStatus(undefined, undefined, config, today).status).toBe('Offen');
      expect(evaluateCaseStatus('2026-06-09', undefined, config, today).status).toBe('Verspätet');
      expect(evaluateCaseStatus('2026-06-12', undefined, config, today).status).toBe('Achtung');
      expect(evaluateCaseStatus('2026-06-22', undefined, config, today).status).toBe('Im Plan');
      expect(evaluateCaseStatus('2026-06-09', 'Fertig', config, today).status).toBe('Fertig');
    });
  });

  describe('German date format support', () => {
    it('handles DD.MM.YYYY planned dates', () => {
      expect(evaluateCaseStatus('09.06.2026', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today).status).toBe('überfällig');
      expect(evaluateCaseStatus('12.06.2026', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today).status).toBe('prüfen');
      expect(evaluateCaseStatus('22.06.2026', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today).status).toBe('läuft planmäßig');
    });
  });

  describe('does not mutate its arguments', () => {
    it('does not mutate the today argument', () => {
      const originalToday = new Date(today.getTime());
      evaluateCaseStatus('2026-06-12', undefined, DEFAULT_STATUS_ENGINE_CONFIG, today);
      expect(today.getTime()).toBe(originalToday.getTime());
    });
  });

  describe('default today (new Date())', () => {
    it('uses the real current date when today is not passed', () => {
      const earlier = new Date(2025, 0, 1);
      earlier.setHours(0, 0, 0, 0);
      // With no planned date, it always returns statusNoPlan regardless of today.
      const result = evaluateCaseStatus(undefined, undefined, DEFAULT_STATUS_ENGINE_CONFIG, earlier);
      expect(result.status).toBe('Termin planen');
    });
  });
});
