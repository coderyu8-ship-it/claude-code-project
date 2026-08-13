/**
 * Money arithmetic over integer minor units (yen, cents, ...).
 *
 * Floating point is never used to hold an amount — only to derive one, and every
 * derivation is immediately rounded back to an integer.
 */

export type Currency = 'JPY' | 'USD' | 'EUR';

export interface Money {
  readonly amount: number;
  readonly currency: Currency;
}

/** Digits after the decimal point in each currency's major unit. */
const MINOR_UNIT_EXPONENT: Record<Currency, number> = {
  JPY: 0,
  USD: 2,
  EUR: 2,
};

export class CurrencyMismatchError extends Error {
  constructor(
    readonly left: Currency,
    readonly right: Currency,
  ) {
    super(`currency mismatch: ${left} and ${right} cannot be combined`);
    this.name = 'CurrencyMismatchError';
  }
}

export type RoundingMode = 'half-up' | 'down' | 'up';

export function money(amount: number, currency: Currency): Money {
  if (!Number.isInteger(amount)) {
    throw new TypeError(`amount must be an integer in minor units, received ${amount}`);
  }
  if (!Number.isSafeInteger(amount)) {
    throw new RangeError(`amount ${amount} exceeds the safe integer range`);
  }
  return { amount, currency };
}

export function zero(currency: Currency): Money {
  return { amount: 0, currency };
}

function assertSameCurrency(left: Money, right: Money): void {
  if (left.currency !== right.currency) {
    throw new CurrencyMismatchError(left.currency, right.currency);
  }
}

export function add(left: Money, right: Money): Money {
  assertSameCurrency(left, right);
  return money(left.amount + right.amount, left.currency);
}

export function subtract(left: Money, right: Money): Money {
  assertSameCurrency(left, right);
  return money(left.amount - right.amount, left.currency);
}

function round(value: number, mode: RoundingMode): number {
  switch (mode) {
    case 'down':
      return Math.trunc(value);
    case 'up':
      return value < 0 ? Math.floor(value) : Math.ceil(value);
    case 'half-up':
      // Math.round breaks ties toward +Infinity, so -2.5 would become -2.
      // Mirroring across zero keeps half-up symmetric for negative amounts.
      return value < 0 ? -Math.round(-value) : Math.round(value);
  }
}

export function multiply(value: Money, factor: number, mode: RoundingMode = 'half-up'): Money {
  if (!Number.isFinite(factor)) {
    throw new TypeError(`factor must be a finite number, received ${factor}`);
  }
  return money(round(value.amount * factor, mode), value.currency);
}

/** Negative when `left` is worth less than `right`, zero when equal. */
export function compare(left: Money, right: Money): number {
  assertSameCurrency(left, right);
  return Math.sign(left.amount - right.amount);
}

export function equals(left: Money, right: Money): boolean {
  return left.currency === right.currency && left.amount === right.amount;
}

export function isZero(value: Money): boolean {
  return value.amount === 0;
}

export function isNegative(value: Money): boolean {
  return value.amount < 0;
}

export function min(left: Money, right: Money): Money {
  return compare(left, right) <= 0 ? left : right;
}

export function max(left: Money, right: Money): Money {
  return compare(left, right) >= 0 ? left : right;
}

/**
 * Split an amount across `ratios` without losing or inventing a single minor unit.
 *
 * Remainders go to the largest fractional shares first (largest-remainder method),
 * so the parts always add back up to the original.
 */
export function allocate(value: Money, ratios: readonly number[]): Money[] {
  if (isNegative(value)) {
    throw new RangeError('cannot allocate a negative amount');
  }
  if (ratios.length === 0) {
    throw new RangeError('ratios must not be empty');
  }
  if (ratios.some((ratio) => !Number.isFinite(ratio) || ratio < 0)) {
    throw new RangeError('every ratio must be a finite, non-negative number');
  }

  const total = ratios.reduce((sum, ratio) => sum + ratio, 0);
  if (total <= 0) {
    throw new RangeError('ratios must sum to a positive number');
  }

  const exact = ratios.map((ratio) => (value.amount * ratio) / total);
  const floors = exact.map((share) => Math.floor(share));
  // Each share loses less than one minor unit to flooring, so the remainder is
  // always smaller than the number of shares — one extra unit each, at most.
  const remainder = value.amount - floors.reduce((sum, share) => sum + share, 0);

  const getsExtraUnit = new Set(
    exact
      .map((share, index) => ({ index, fraction: share - Math.floor(share) }))
      .sort((a, b) => b.fraction - a.fraction || a.index - b.index)
      .slice(0, remainder)
      .map((entry) => entry.index),
  );

  return floors.map((share, index) =>
    money(share + (getsExtraUnit.has(index) ? 1 : 0), value.currency),
  );
}

export function format(value: Money, locale = 'ja-JP'): string {
  const exponent = MINOR_UNIT_EXPONENT[value.currency];
  return new Intl.NumberFormat(locale, {
    style: 'currency',
    currency: value.currency,
    minimumFractionDigits: exponent,
    maximumFractionDigits: exponent,
  }).format(value.amount / 10 ** exponent);
}
