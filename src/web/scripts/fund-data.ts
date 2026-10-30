// Fictional fund-administration data for the "fund-admin" scenario. Every name here is invented.

export const FUND = 'Halverson Growth Partners III, L.P.';
export const GP = 'Halverson GP III, LLC';
export const ADMIN = 'Pinecrest Fund Services';

export interface LP { id: string; name: string; type: string; commitmentM: number; gp?: boolean }

/** commitments in USD millions; total 250 */
export const LPS: LP[] = [
  { id: 'LP-001', name: 'Brightwater County Employees\' Retirement System', type: 'Public pension', commitmentM: 40 },
  { id: 'LP-002', name: 'Calloway Family Foundation', type: 'Foundation', commitmentM: 15 },
  { id: 'LP-003', name: 'Northgate University Endowment', type: 'Endowment', commitmentM: 30 },
  { id: 'LP-004', name: 'Meridale Mutual Insurance Company', type: 'Insurance', commitmentM: 25 },
  { id: 'LP-005', name: 'Ashford Point Pension Trust', type: 'Corporate pension', commitmentM: 35 },
  { id: 'LP-006', name: 'Redstone Capital Holdings LLC', type: 'Corporate', commitmentM: 20 },
  { id: 'LP-007', name: 'Larkspur Fund of Funds IV, L.P.', type: 'Fund of funds', commitmentM: 27.5 },
  { id: 'LP-008', name: 'Ellison & Vance Family Office', type: 'Family office', commitmentM: 10 },
  { id: 'LP-009', name: 'Tidewell Health System Retirement Plan', type: 'Corporate pension', commitmentM: 22.5 },
  { id: 'LP-010', name: 'Quarry Hill Partners LLC', type: 'Corporate', commitmentM: 12.5 },
  { id: 'LP-011', name: 'Osprey Ridge Charitable Trust', type: 'Foundation', commitmentM: 10 },
  { id: 'LP-012', name: GP, type: 'General partner', commitmentM: 2.5, gp: true },
];

export const TOTAL_M = LPS.reduce((s, l) => s + l.commitmentM, 0); // 250
export const M = 1_000_000;

/** capital calls 1-3 as a fraction of commitments */
export const PRIOR_CALLS = [0.12, 0.10, 0.08];

/** Capital Call No. 4 */
export const CALL4 = {
  notice: [2026, 10, 15] as const,
  due: [2026, 10, 30] as const,
  items: [
    { label: 'Investment — Alder Robotics, Inc. Series C follow-on', amount: 14_500_000 },
    { label: 'Investment — Juniper Lane Logistics, Inc. bridge note', amount: 1_250_000 },
    { label: 'Partnership expenses (audit, tax, administration)', amount: 750_000 },
  ],
};
export const CALL4_TOTAL = CALL4.items.reduce((s, i) => s + i.amount, 0); // 16,500,000

/** Q4 2026 management fee: 2% a year on fee-paying commitments (the GP pays none), actual/365 */
export const FEE = { rate: 0.02, start: [2026, 10, 1] as const, end: [2026, 12, 31] as const, basis: 365 };
export const FEE_DAYS = 92;
export const feeFor = (lp: LP) => (lp.gp ? 0 : lp.commitmentM * M * FEE.rate * FEE_DAYS / FEE.basis);

/** Q3 2026 capital account roll-forward, per USD 1m of commitment */
export const Q3 = {
  beginningPerM: 231_750,
  contributionsPerM: 0.08 * M,   // Capital Call No. 3 (8%)
  distributionsPerM: 5_000,      // 1,250,000 distribution / 250m
  feePerM: 0.02 / 4 * M,         // 0.5% a quarter, fee-paying LPs only
  realizedPerM: 4_200,
  unrealizedPerM: 18_400,
};
export function rollForward(lp: LP) {
  const k = lp.commitmentM;
  const b = k * Q3.beginningPerM, c = k * Q3.contributionsPerM, d = k * Q3.distributionsPerM;
  const f = lp.gp ? 0 : k * Q3.feePerM, g = k * Q3.realizedPerM, h = k * Q3.unrealizedPerM;
  return { b, c, d, f, g, h, end: b + c - d - f + g + h };
}
