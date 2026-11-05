import {
  aggregateShippedPlaces,
  placesForRibbon,
} from '@/components/home/shipped-places';

describe('aggregateShippedPlaces', () => {
  test('folds Ohio aliases into one state and leaves the United States off the country list', () => {
    const places = aggregateShippedPlaces([
      { country: 'USA', state: 'OH' },
      { country: 'United States', state: 'Ohio' },
      { country: 'u.s.a.', state: ' ohio ' },
      { country: '', state: 'Ohio' },
    ]);

    expect(places.countries).toEqual([]);
    expect(places.countryCount).toBe(0);
    expect(places.usStates).toEqual([{ name: 'Ohio', orderCount: 4 }]);
    expect(places.usStateCount).toBe(1);
  });

  test('counts Puerto Rico from its name and its abbreviation', () => {
    const places = aggregateShippedPlaces([
      { country: 'US', state: 'Puerto Rico' },
      { country: 'United States of America', state: 'PR' },
    ]);

    expect(places.usStates).toEqual([{ name: 'Puerto Rico', orderCount: 2 }]);
    expect(places.countries).toEqual([]);
  });

  test('counts Canada and ignores its province', () => {
    const places = aggregateShippedPlaces([
      { country: 'Canada', state: 'Ontario' },
    ]);

    expect(places.countries).toEqual([{ name: 'Canada', orderCount: 1 }]);
    expect(places.usStates).toEqual([]);
  });

  test('maps UK and Holland onto the listed country names', () => {
    const places = aggregateShippedPlaces([
      { country: 'UK', state: '' },
      { country: 'u.k.', state: 'London' },
      { country: 'Great Britain', state: '' },
      { country: 'Holland', state: '' },
      { country: 'the netherlands', state: '' },
    ]);

    expect(places.countries).toEqual([
      { name: 'United Kingdom', orderCount: 3 },
      { name: 'The Netherlands', orderCount: 2 },
    ]);
    expect(places.usStates).toEqual([]);
  });

  test('drops Italy and Washington, D.C.', () => {
    const places = aggregateShippedPlaces([
      { country: 'Italy', state: 'Rome' },
      { country: 'United States', state: 'Washington, D.C.' },
      { country: 'USA', state: 'DC' },
    ]);

    expect(places).toEqual({
      countries: [],
      usStates: [],
      countryCount: 0,
      usStateCount: 0,
    });
  });

  test('sorts by order count, then by name', () => {
    const places = aggregateShippedPlaces([
      { country: 'Mexico', state: '' },
      { country: 'Mexico', state: '' },
      { country: 'Australia', state: '' },
      { country: 'Canada', state: '' },
      { country: 'US', state: 'Texas' },
      { country: 'US', state: 'OH' },
      { country: 'US', state: 'Ohio' },
      { country: 'US', state: 'AL' },
      { country: 'US', state: 'Alabama' },
    ]);

    expect(places.countries).toEqual([
      { name: 'Mexico', orderCount: 2 },
      { name: 'Australia', orderCount: 1 },
      { name: 'Canada', orderCount: 1 },
    ]);
    expect(places.usStates).toEqual([
      { name: 'Alabama', orderCount: 2 },
      { name: 'Ohio', orderCount: 2 },
      { name: 'Texas', orderCount: 1 },
    ]);
    expect(places.countryCount).toBe(3);
    expect(places.usStateCount).toBe(3);
  });
});

describe('placesForRibbon', () => {
  test('keeps every listed name and sorts counted places first', () => {
    expect(
      placesForRibbon(
        ['Canada', 'Australia', 'Mexico'],
        [
          { name: 'Mexico', orderCount: 2 },
          { name: 'Canada', orderCount: 1 },
        ],
      ),
    ).toEqual([
      { name: 'Mexico', orderCount: 2 },
      { name: 'Canada', orderCount: 1 },
      { name: 'Australia', orderCount: 0 },
    ]);
  });

  test('keeps listed order when counts are equal', () => {
    expect(
      placesForRibbon(['Canada', 'Australia', 'Mexico'], []).map(
        (place) => place.name,
      ),
    ).toEqual(['Canada', 'Australia', 'Mexico']);
  });
});
