import { isoToUsDate, isValidUsDate, usDateToIso } from 'lib/helpers/us-date';

describe('us date', () => {
  test('round-trips an ISO day and a US date', () => {
    expect(isoToUsDate('1977-02-04')).toBe('02/04/1977');
    expect(usDateToIso('02/04/1977')).toBe('1977-02-04');
    expect(isoToUsDate(usDateToIso('02/04/1977'))).toBe('02/04/1977');
    expect(usDateToIso(isoToUsDate('1977-02-04'))).toBe('1977-02-04');
  });

  test('rejects an incomplete or impossible date', () => {
    expect(isValidUsDate('02/04/1977')).toBe(true);
    expect(isValidUsDate('02/29/1976')).toBe(true);
    expect(isValidUsDate('13/40/19')).toBe(false);
    expect(isValidUsDate('02/30/1977')).toBe(false);
    expect(isValidUsDate('02/29/1977')).toBe(false);
    expect(usDateToIso('02/30/1977')).toBeNull();
    expect(usDateToIso('13/40/19')).toBeNull();
    expect(usDateToIso('')).toBeNull();
    expect(isoToUsDate(null)).toBe('');
    expect(isoToUsDate('1977-02-30')).toBe('');
  });
});
