import { describe, expect, it } from 'vitest';

import { type Discount, discountAmount, totalDiscount } from '../src/core/discount.js';
import { CurrencyMismatchError, money, zero } from '../src/core/money.js';

describe('discountAmount — percentage', () => {
  it('takes the given percentage off the base', () => {
    expect(discountAmount(money(1000, 'JPY'), { kind: 'percentage', percent: 20 })).toEqual(
      money(200, 'JPY'),
    );
  });

  it('rounds half-up on uneven percentages', () => {
    expect(discountAmount(money(101, 'JPY'), { kind: 'percentage', percent: 50 })).toEqual(
      money(51, 'JPY'),
    );
  });

  it('takes nothing at 0% and everything at 100%', () => {
    expect(discountAmount(money(1000, 'JPY'), { kind: 'percentage', percent: 0 })).toEqual(
      zero('JPY'),
    );
    expect(discountAmount(money(1000, 'JPY'), { kind: 'percentage', percent: 100 })).toEqual(
      money(1000, 'JPY'),
    );
  });

  it('caps at maxAmount when the percentage exceeds it', () => {
    const discount: Discount = {
      kind: 'percentage',
      percent: 50,
      maxAmount: money(300, 'JPY'),
    };
    expect(discountAmount(money(1000, 'JPY'), discount)).toEqual(money(300, 'JPY'));
  });

  it('leaves the percentage alone when it stays under maxAmount', () => {
    const discount: Discount = {
      kind: 'percentage',
      percent: 10,
      maxAmount: money(300, 'JPY'),
    };
    expect(discountAmount(money(1000, 'JPY'), discount)).toEqual(money(100, 'JPY'));
  });

  it('rejects a percentage outside 0–100', () => {
    expect(() => discountAmount(money(100, 'JPY'), { kind: 'percentage', percent: -1 })).toThrow(
      RangeError,
    );
    expect(() => discountAmount(money(100, 'JPY'), { kind: 'percentage', percent: 101 })).toThrow(
      RangeError,
    );
    expect(() =>
      discountAmount(money(100, 'JPY'), { kind: 'percentage', percent: Number.NaN }),
    ).toThrow(RangeError);
  });

  it('rejects a maxAmount in another currency', () => {
    expect(() =>
      discountAmount(money(1000, 'JPY'), {
        kind: 'percentage',
        percent: 50,
        maxAmount: money(300, 'USD'),
      }),
    ).toThrow(CurrencyMismatchError);
  });

  it('rejects a negative maxAmount', () => {
    expect(() =>
      discountAmount(money(1000, 'JPY'), {
        kind: 'percentage',
        percent: 50,
        maxAmount: money(-1, 'JPY'),
      }),
    ).toThrow(/maxAmount must not be negative/);
  });
});

describe('discountAmount — fixed', () => {
  it('takes the fixed amount off', () => {
    expect(discountAmount(money(1000, 'JPY'), { kind: 'fixed', amount: money(250, 'JPY') })).toEqual(
      money(250, 'JPY'),
    );
  });

  it('never takes more than the base', () => {
    expect(
      discountAmount(money(200, 'JPY'), { kind: 'fixed', amount: money(1000, 'JPY') }),
    ).toEqual(money(200, 'JPY'));
  });

  it('rejects a discount in another currency', () => {
    expect(() =>
      discountAmount(money(1000, 'JPY'), { kind: 'fixed', amount: money(250, 'USD') }),
    ).toThrow(CurrencyMismatchError);
  });

  it('rejects a negative discount', () => {
    expect(() =>
      discountAmount(money(1000, 'JPY'), { kind: 'fixed', amount: money(-250, 'JPY') }),
    ).toThrow(/must not be negative/);
  });
});

describe('discountAmount — shared guards', () => {
  it('rejects a negative base', () => {
    expect(() =>
      discountAmount(money(-1, 'JPY'), { kind: 'fixed', amount: money(1, 'JPY') }),
    ).toThrow(/negative base/);
  });
});

describe('totalDiscount', () => {
  it('returns zero for an empty discount list', () => {
    expect(totalDiscount(money(1000, 'JPY'), [])).toEqual(zero('JPY'));
  });

  it('applies each discount to what the previous one left behind', () => {
    // 50% of 1000 = 500, then 50% of the remaining 500 = 250.
    const discounts: Discount[] = [
      { kind: 'percentage', percent: 50 },
      { kind: 'percentage', percent: 50 },
    ];
    expect(totalDiscount(money(1000, 'JPY'), discounts)).toEqual(money(750, 'JPY'));
  });

  it('mixes fixed and percentage discounts in order', () => {
    // 200 off 1000 leaves 800, then 10% of 800 = 80.
    const discounts: Discount[] = [
      { kind: 'fixed', amount: money(200, 'JPY') },
      { kind: 'percentage', percent: 10 },
    ];
    expect(totalDiscount(money(1000, 'JPY'), discounts)).toEqual(money(280, 'JPY'));
  });

  it('stops at the base amount rather than going past zero', () => {
    const discounts: Discount[] = [
      { kind: 'fixed', amount: money(800, 'JPY') },
      { kind: 'fixed', amount: money(800, 'JPY') },
    ];
    expect(totalDiscount(money(1000, 'JPY'), discounts)).toEqual(money(1000, 'JPY'));
  });
});
