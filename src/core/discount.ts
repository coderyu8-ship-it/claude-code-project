import {
  type Money,
  CurrencyMismatchError,
  isNegative,
  min,
  multiply,
  zero,
} from './money.js';

export interface PercentageDiscount {
  readonly kind: 'percentage';
  readonly percent: number;
  /** Optional ceiling, e.g. "20% off, up to ¥1,000". */
  readonly maxAmount?: Money;
}

export interface FixedDiscount {
  readonly kind: 'fixed';
  readonly amount: Money;
}

export type Discount = PercentageDiscount | FixedDiscount;

/**
 * The amount a single discount takes off `base`.
 *
 * Never exceeds `base` and never goes negative, so a stack of discounts can only
 * drive a total down to zero.
 */
export function discountAmount(base: Money, discount: Discount): Money {
  if (isNegative(base)) {
    throw new RangeError('cannot discount a negative base amount');
  }

  switch (discount.kind) {
    case 'percentage': {
      if (!Number.isFinite(discount.percent) || discount.percent < 0 || discount.percent > 100) {
        throw new RangeError(`percent must be between 0 and 100, received ${discount.percent}`);
      }
      const raw = multiply(base, discount.percent / 100);
      const capped = discount.maxAmount === undefined ? raw : capBy(raw, discount.maxAmount);
      return min(capped, base);
    }
    case 'fixed': {
      if (discount.amount.currency !== base.currency) {
        throw new CurrencyMismatchError(base.currency, discount.amount.currency);
      }
      if (isNegative(discount.amount)) {
        throw new RangeError('a fixed discount must not be negative');
      }
      return min(discount.amount, base);
    }
  }
}

function capBy(value: Money, ceiling: Money): Money {
  if (ceiling.currency !== value.currency) {
    throw new CurrencyMismatchError(value.currency, ceiling.currency);
  }
  if (isNegative(ceiling)) {
    throw new RangeError('maxAmount must not be negative');
  }
  return min(value, ceiling);
}

/**
 * Apply discounts in order, each one computed against the amount left by the
 * previous. Returns the combined reduction, never more than `base`.
 */
export function totalDiscount(base: Money, discounts: readonly Discount[]): Money {
  let remaining = base;
  let taken = zero(base.currency);

  for (const discount of discounts) {
    const step = discountAmount(remaining, discount);
    taken = { amount: taken.amount + step.amount, currency: taken.currency };
    remaining = { amount: remaining.amount - step.amount, currency: remaining.currency };
  }

  return taken;
}
