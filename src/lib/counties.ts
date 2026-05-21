/**
 * SC + NC county catalog for the cahpCounty Property field.
 *
 * `cahpCounty` is stored as a single string on SharePoint — when a property
 * spans more than one county (e.g., a single-family home on a parcel that
 * straddles two), the values are joined with ", " so existing list views
 * still render something sensible. The app parses on read and shows chips.
 */

export const SC_COUNTIES = [
  'Abbeville', 'Aiken', 'Allendale', 'Anderson', 'Bamberg', 'Barnwell',
  'Beaufort', 'Berkeley', 'Calhoun', 'Charleston', 'Cherokee', 'Chester',
  'Chesterfield', 'Clarendon', 'Colleton', 'Darlington', 'Dillon', 'Dorchester',
  'Edgefield', 'Fairfield', 'Florence', 'Georgetown', 'Greenville', 'Greenwood',
  'Hampton', 'Horry', 'Jasper', 'Kershaw', 'Lancaster', 'Laurens', 'Lee',
  'Lexington', 'Marion', 'Marlboro', 'McCormick', 'Newberry', 'Oconee',
  'Orangeburg', 'Pickens', 'Richland', 'Saluda', 'Spartanburg', 'Sumter',
  'Union', 'Williamsburg', 'York',
] as const;

export const NC_COUNTIES = [
  'Alamance', 'Alexander', 'Alleghany', 'Anson', 'Ashe', 'Avery', 'Beaufort',
  'Bertie', 'Bladen', 'Brunswick', 'Buncombe', 'Burke', 'Cabarrus', 'Caldwell',
  'Camden', 'Carteret', 'Caswell', 'Catawba', 'Chatham', 'Cherokee', 'Chowan',
  'Clay', 'Cleveland', 'Columbus', 'Craven', 'Cumberland', 'Currituck', 'Dare',
  'Davidson', 'Davie', 'Duplin', 'Durham', 'Edgecombe', 'Forsyth', 'Franklin',
  'Gaston', 'Gates', 'Graham', 'Granville', 'Greene', 'Guilford', 'Halifax',
  'Harnett', 'Haywood', 'Henderson', 'Hertford', 'Hoke', 'Hyde', 'Iredell',
  'Jackson', 'Johnston', 'Jones', 'Lee', 'Lenoir', 'Lincoln', 'Macon',
  'Madison', 'Martin', 'McDowell', 'Mecklenburg', 'Mitchell', 'Montgomery',
  'Moore', 'Nash', 'New Hanover', 'Northampton', 'Onslow', 'Orange', 'Pamlico',
  'Pasquotank', 'Pender', 'Perquimans', 'Person', 'Pitt', 'Polk', 'Randolph',
  'Richmond', 'Robeson', 'Rockingham', 'Rowan', 'Rutherford', 'Sampson',
  'Scotland', 'Stanly', 'Stokes', 'Surry', 'Swain', 'Transylvania', 'Tyrrell',
  'Union', 'Vance', 'Wake', 'Warren', 'Washington', 'Watauga', 'Wayne',
  'Wilkes', 'Wilson', 'Yadkin', 'Yancey',
] as const;

/**
 * Full catalog with state suffix for disambiguation — e.g., "Beaufort (SC)" vs
 * "Beaufort (NC)", "Cherokee (SC)" vs "Cherokee (NC)", "Lee (SC)" vs "Lee (NC)",
 * "Union (SC)" vs "Union (NC)". This is the canonical value stored on the
 * Property record.
 */
export const ALL_COUNTIES: string[] = [
  ...SC_COUNTIES.map((c) => `${c} (SC)`),
  ...NC_COUNTIES.map((c) => `${c} (NC)`),
  'Other',
].sort((a, b) => a.localeCompare(b));

/**
 * Parse the stored cahpCounty value into an array. Tolerates legacy single
 * values and "Other" sentinel; ignores empty pieces.
 */
export function parseCounties(value: string | undefined | null): string[] {
  if (!value) return [];
  return value
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

/** Inverse of parseCounties — join an array for storage. */
export function joinCounties(counties: string[]): string {
  return counties.join(', ');
}
