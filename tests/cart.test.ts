import { describe, expect, it } from 'vitest';

import {
  type Cart,
  addLine,
  calculateTotals,
  emptyCart,
  isEmpty,
  itemCount,
  line,
  lineTotal,
  removeLine,
  subtotal,
} from '../src/core/cart.js';
import type { Discount } from '../src/core/discount.js';
import { CurrencyMismatchError, money, zero } from '../src/core/money.js';

const jpyCart = (lines: Cart['lines'], rest: Partial<Cart> = {}): Cart => ({
  currency: 'JPY',
  lines,
  ...rest,
});

describe('lineTotal', () => {
  it('multiplies unit price by quantity', () => {
    expect(lineTotal(line('SKU-1', 500, 3, 'JPY'))).toEqual(money(1500, 'JPY'));
  });
});

describe('subtotal', () => {
  it('sums every line', () => {
    const cart = jpyCart([line('SKU-1', 500, 3, 'JPY'), line('SKU-2', 1200, 1, 'JPY')]);
    expect(subtotal(cart)).toEqual(money(2700, 'JPY'));
  });

  it('is zero for an empty cart', () => {
    expect(subtotal(emptyCart('JPY'))).toEqual(zero('JPY'));
  });

  it('rejects a line priced in another currency', () => {
    const cart = jpyCart([line('SKU-1', 500, 1, 'USD')]);
    expect(() => subtotal(cart)).toThrow(CurrencyMismatchError);
  });

  it('rejects a negative unit price', () => {
    const cart = jpyCart([{ sku: 'SKU-1', unitPrice: money(-500, 'JPY'), quantity: 1 }]);
    expect(() => subtotal(cart)).toThrow(/must not be negative/);
  });

  it('rejects a zero, negative, or fractional quantity', () => {
    expect(() => subtotal(jpyCart([line('SKU-1', 500, 0, 'JPY')]))).toThrow(/positive integer/);
    expect(() => subtotal(jpyCart([line('SKU-1', 500, -2, 'JPY')]))).toThrow(/positive integer/);
    expect(() => subtotal(jpyCart([line('SKU-1', 500, 1.5, 'JPY')]))).toThrow(/positive integer/);
  });
});

describe('calculateTotals', () => {
  it('returns a subtotal-only breakdown with no discounts or tax', () => {
    const totals = calculateTotals(jpyCart([line('SKU-1', 1000, 2, 'JPY')]));
    expect(totals).toEqual({
      subtotal: money(2000, 'JPY'),
      discountTotal: zero('JPY'),
      taxableBase: money(2000, 'JPY'),
      tax: zero('JPY'),
      total: money(2000, 'JPY'),
    });
  });

  it('taxes the post-discount amount, not the subtotal', () => {
    const discounts: Discount[] = [{ kind: 'percentage', percent: 10 }];
    const totals = calculateTotals(
      jpyCart([line('SKU-1', 1000, 1, 'JPY')], { discounts, taxRatePercent: 10 }),
    );

    expect(totals.subtotal).toEqual(money(1000, 'JPY'));
    expect(totals.discountTotal).toEqual(money(100, 'JPY'));
    expect(totals.taxableBase).toEqual(money(900, 'JPY'));
    expect(totals.tax).toEqual(money(90, 'JPY'));
    expect(totals.total).toEqual(money(990, 'JPY'));
  });

  it('rounds tax half-up', () => {
    // 10% of 1005 is 100.5 -> 101.
    const totals = calculateTotals(jpyCart([line('SKU-1', 1005, 1, 'JPY')], { taxRatePercent: 10 }));
    expect(totals.tax).toEqual(money(101, 'JPY'));
  });

  it('bottoms out at zero when discounts exceed the subtotal', () => {
    const discounts: Discount[] = [{ kind: 'fixed', amount: money(5000, 'JPY') }];
    const totals = calculateTotals(
      jpyCart([line('SKU-1', 1000, 1, 'JPY')], { discounts, taxRatePercent: 10 }),
    );

    expect(totals.taxableBase).toEqual(zero('JPY'));
    expect(totals.tax).toEqual(zero('JPY'));
    expect(totals.total).toEqual(zero('JPY'));
  });

  it('rejects a negative or non-finite tax rate', () => {
    expect(() =>
      calculateTotals(jpyCart([line('SKU-1', 100, 1, 'JPY')], { taxRatePercent: -1 })),
    ).toThrow(/non-negative/);
    expect(() =>
      calculateTotals(jpyCart([line('SKU-1', 100, 1, 'JPY')], { taxRatePercent: Number.NaN })),
    ).toThrow(/non-negative/);
  });
});

describe('cart mutation helpers', () => {
  it('starts empty', () => {
    const cart = emptyCart('JPY');
    expect(isEmpty(cart)).toBe(true);
    expect(itemCount(cart)).toBe(0);
  });

  it('appends a new SKU', () => {
    const cart = addLine(emptyCart('JPY'), line('SKU-1', 500, 2, 'JPY'));
    expect(cart.lines).toHaveLength(1);
    expect(isEmpty(cart)).toBe(false);
    expect(itemCount(cart)).toBe(2);
  });

  it('merges quantities for a SKU already in the cart', () => {
    const cart = addLine(
      addLine(emptyCart('JPY'), line('SKU-1', 500, 2, 'JPY')),
      line('SKU-1', 500, 3, 'JPY'),
    );

    expect(cart.lines).toHaveLength(1);
    expect(cart.lines[0]?.quantity).toBe(5);
    expect(itemCount(cart)).toBe(5);
  });

  it('leaves other SKUs untouched when merging', () => {
    const base = addLine(
      addLine(emptyCart('JPY'), line('SKU-1', 500, 1, 'JPY')),
      line('SKU-2', 300, 4, 'JPY'),
    );
    const merged = addLine(base, line('SKU-1', 500, 1, 'JPY'));

    expect(merged.lines).toHaveLength(2);
    expect(merged.lines[0]?.quantity).toBe(2);
    expect(merged.lines[1]?.quantity).toBe(4);
  });

  it('refuses to merge the same SKU at a different price', () => {
    const cart = addLine(emptyCart('JPY'), line('SKU-1', 500, 1, 'JPY'));
    expect(() => addLine(cart, line('SKU-1', 900, 1, 'JPY'))).toThrow(/price conflict/);
  });

  it('validates the line being added', () => {
    expect(() => addLine(emptyCart('JPY'), line('SKU-1', 500, 0, 'JPY'))).toThrow(
      /positive integer/,
    );
    expect(() => addLine(emptyCart('JPY'), line('SKU-1', 500, 1, 'USD'))).toThrow(
      CurrencyMismatchError,
    );
  });

  it('does not mutate the cart it was given', () => {
    const original = emptyCart('JPY');
    addLine(original, line('SKU-1', 500, 1, 'JPY'));
    expect(original.lines).toHaveLength(0);
  });

  it('removes a SKU and ignores an unknown one', () => {
    const cart = addLine(emptyCart('JPY'), line('SKU-1', 500, 1, 'JPY'));
    expect(removeLine(cart, 'SKU-1').lines).toHaveLength(0);
    expect(removeLine(cart, 'SKU-404').lines).toHaveLength(1);
  });
});
