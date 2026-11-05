export type ShipmentAddress = {
  country: string;
  state: string;
};

export type ShippedPlace = {
  name: string;
  orderCount: number;
};

export type ShippedPlaces = {
  countries: ShippedPlace[];
  usStates: ShippedPlace[];
  countryCount: number;
  usStateCount: number;
};

const COUNTRIES = [
  'Canada',
  'Australia',
  'United Kingdom',
  'Germany',
  'Norway',
  'Switzerland',
  'France',
  'Ireland',
  'Finland',
  'Spain',
  'The Netherlands',
  'Belgium',
  'Brazil',
  'China',
  'Denmark',
  'Greece',
  'Japan',
  'Malaysia',
  'Mexico',
  'South Africa',
] satisfies readonly string[];

const STATES = [
  ['AL', 'Alabama'],
  ['AK', 'Alaska'],
  ['AZ', 'Arizona'],
  ['AR', 'Arkansas'],
  ['CA', 'California'],
  ['CO', 'Colorado'],
  ['CT', 'Connecticut'],
  ['DE', 'Delaware'],
  ['FL', 'Florida'],
  ['GA', 'Georgia'],
  ['HI', 'Hawaii'],
  ['ID', 'Idaho'],
  ['IL', 'Illinois'],
  ['IN', 'Indiana'],
  ['IA', 'Iowa'],
  ['KS', 'Kansas'],
  ['KY', 'Kentucky'],
  ['LA', 'Louisiana'],
  ['ME', 'Maine'],
  ['MD', 'Maryland'],
  ['MA', 'Massachusetts'],
  ['MI', 'Michigan'],
  ['MN', 'Minnesota'],
  ['MS', 'Mississippi'],
  ['MO', 'Missouri'],
  ['MT', 'Montana'],
  ['NE', 'Nebraska'],
  ['NV', 'Nevada'],
  ['NH', 'New Hampshire'],
  ['NJ', 'New Jersey'],
  ['NM', 'New Mexico'],
  ['NY', 'New York'],
  ['NC', 'North Carolina'],
  ['ND', 'North Dakota'],
  ['OH', 'Ohio'],
  ['OK', 'Oklahoma'],
  ['OR', 'Oregon'],
  ['PA', 'Pennsylvania'],
  ['RI', 'Rhode Island'],
  ['SC', 'South Carolina'],
  ['SD', 'South Dakota'],
  ['TN', 'Tennessee'],
  ['TX', 'Texas'],
  ['UT', 'Utah'],
  ['VT', 'Vermont'],
  ['VA', 'Virginia'],
  ['WA', 'Washington'],
  ['WV', 'West Virginia'],
  ['WI', 'Wisconsin'],
  ['WY', 'Wyoming'],
  ['PR', 'Puerto Rico'],
] satisfies readonly (readonly [string, string])[];

const US_COUNTRY_KEYS = new Set([
  'us',
  'usa',
  'united states',
  'united states of america',
]);

const countryByKey = new Map<string, string>(
  COUNTRIES.map((name) => [normalizePlace(name), name]),
);
countryByKey.set('uk', 'United Kingdom');
countryByKey.set('great britain', 'United Kingdom');
countryByKey.set('holland', 'The Netherlands');
countryByKey.set('netherlands', 'The Netherlands');

const stateByKey = new Map<string, string>();
for (const [abbreviation, name] of STATES) {
  stateByKey.set(normalizePlace(name), name);
  stateByKey.set(normalizePlace(abbreviation), name);
}

function normalizePlace(value: string) {
  return value.trim().toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ');
}

function addCount(counts: Map<string, number>, name: string) {
  counts.set(name, (counts.get(name) ?? 0) + 1);
}

function toShippedPlaces(counts: Map<string, number>) {
  return [...counts]
    .map(([name, orderCount]) => ({ name, orderCount }))
    .toSorted(
      (left, right) =>
        right.orderCount - left.orderCount ||
        left.name.localeCompare(right.name),
    );
}

export function aggregateShippedPlaces(
  orders: readonly ShipmentAddress[],
): ShippedPlaces {
  const countryCounts = new Map<string, number>();
  const stateCounts = new Map<string, number>();

  for (const order of orders) {
    const countryKey = normalizePlace(order.country);
    const stateKey = normalizePlace(order.state);
    if (US_COUNTRY_KEYS.has(countryKey) || countryKey === '') {
      const stateName = stateByKey.get(stateKey);
      if (stateName) addCount(stateCounts, stateName);
      continue;
    }

    const countryName = countryByKey.get(countryKey);
    if (countryName) addCount(countryCounts, countryName);
  }

  const countries = toShippedPlaces(countryCounts);
  const usStates = toShippedPlaces(stateCounts);

  return {
    countries,
    usStates,
    countryCount: countries.length,
    usStateCount: usStates.length,
  };
}
