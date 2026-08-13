import { describe, expect, it } from 'vitest';

import {
  CurrencyMismatchError,
  add,
  allocate,
  compare,
  equals,
  format,
  isNegative,
  isZero,
  max,
  min,
  money,
  multiply,
  subtract,
  zero,
} from '../src/core/money.js';

describe('money', () => {
  it('builds a value from integer minor units', () => {
    expect(money(1200, 'JPY')).toEqual({ amount: 1200, currency: 'JPY' });
  });

  it('rejects fractional minor units', () => {
    expect(() => money(10.5, 'USD')).toThrow(TypeError);
  });

  it('rejects amounts beyond the safe integer range', () => {
    expect(() => money(2 ** 53, 'USD')).toThrow(RangeError);
  });

  it('builds a zero value', () => {
    expect(zero('EUR')).toEqual({ amount: 0, currency: 'EUR' });
  });
});

describe('add / subtract', () => {
  it('adds two amounts of the same currency', () => {
    expect(add(money(300, 'JPY'), money(250, 'JPY'))).toEqual(money(550, 'JPY'));
  });

  it('subtracts and allows the result to go negative', () => {
    expect(subtract(money(300, 'JPY'), money(500, 'JPY'))).toEqual(money(-200, 'JPY'));
  });

  it('refuses to mix currencies', () => {
    expect(() => add(money(100, 'JPY'), money(100, 'USD'))).toThrow(CurrencyMismatchError);
    expect(() => subtract(money(100, 'JPY'), money(100, 'USD'))).toThrow(
      /currency mismatch: JPY and USD/,
    );
  });
});

describe('multiply', () => {
  it('rounds half-up by default', () => {
    expect(multiply(money(101, 'JPY'), 0.5)).toEqual(money(51, 'JPY'));
  });

  it('rounds half-up symmetrically for negative amounts', () => {
    expect(multiply(money(-101, 'JPY'), 0.5)).toEqual(money(-51, 'JPY'));
  });

  it('truncates toward zero when rounding down', () => {
    expect(multiply(money(199, 'JPY'), 0.5, 'down')).toEqual(money(99, 'JPY'));
    expect(multiply(money(-199, 'JPY'), 0.5, 'down')).toEqual(money(-99, 'JPY'));
  });

  it('rounds away from zero when rounding up', () => {
    expect(multiply(money(101, 'JPY'), 0.5, 'up')).toEqual(money(51, 'JPY'));
    expect(multiply(money(-101, 'JPY'), 0.5, 'up')).toEqual(money(-51, 'JPY'));
  });

  it('rejects a non-finite factor', () => {
    expect(() => multiply(money(100, 'JPY'), Number.NaN)).toThrow(TypeError);
    expect(() => multiply(money(100, 'JPY'), Number.POSITIVE_INFINITY)).toThrow(TypeError);
  });
});

describe('comparison helpers', () => {
  it('orders two amounts', () => {
    expect(compare(money(100, 'JPY'), money(200, 'JPY'))).toBe(-1);
    expect(compare(money(200, 'JPY'), money(100, 'JPY'))).toBe(1);
    expect(compare(money(100, 'JPY'), money(100, 'JPY'))).toBe(0);
  });

  it('refuses to compare across currencies', () => {
    expect(() => compare(money(100, 'JPY'), money(100, 'EUR'))).toThrow(CurrencyMismatchError);
  });

  it('checks equality including the currency', () => {
    expect(equals(money(100, 'JPY'), money(100, 'JPY'))).toBe(true);
    expect(equals(money(100, 'JPY'), money(100, 'USD'))).toBe(false);
    expect(equals(money(100, 'JPY'), money(101, 'JPY'))).toBe(false);
  });

  it('detects zero and negative amounts', () => {
    expect(isZero(zero('JPY'))).toBe(true);
    expect(isZero(money(1, 'JPY'))).toBe(false);
    expect(isNegative(money(-1, 'JPY'))).toBe(true);
    expect(isNegative(zero('JPY'))).toBe(false);
  });

  it('picks the smaller and larger amount', () => {
    const small = money(100, 'JPY');
    const large = money(900, 'JPY');
    expect(min(small, large)).toEqual(small);
    expect(min(large, small)).toEqual(small);
    expect(max(small, large)).toEqual(large);
    expect(max(large, small)).toEqual(large);
  });

  it('returns either side when both are equal', () => {
    expect(min(money(5, 'JPY'), money(5, 'JPY'))).toEqual(money(5, 'JPY'));
    expect(max(money(5, 'JPY'), money(5, 'JPY'))).toEqual(money(5, 'JPY'));
  });
});

describe('allocate', () => {
  it('splits evenly when it divides cleanly', () => {
    expect(allocate(money(900, 'JPY'), [1, 1, 1])).toEqual([
      money(300, 'JPY'),
      money(300, 'JPY'),
      money(300, 'JPY'),
    ]);
  });

  it('hands remainders to the largest fractional shares', () => {
    expect(allocate(money(100, 'JPY'), [1, 1, 1])).toEqual([
      money(34, 'JPY'),
      money(33, 'JPY'),
      money(33, 'JPY'),
    ]);
  });

  it('respects weighted ratios', () => {
    expect(allocate(money(1000, 'JPY'), [7, 3])).toEqual([money(700, 'JPY'), money(300, 'JPY')]);
  });

  it('never loses or invents a minor unit', () => {
    const parts = allocate(money(1_000_001, 'JPY'), [3, 3, 3, 1]);
    const sum = parts.reduce((total, part) => total + part.amount, 0);
    expect(sum).toBe(1_000_001);
  });

  it('tolerates a zero ratio', () => {
    expect(allocate(money(100, 'JPY'), [1, 0])).toEqual([money(100, 'JPY'), zero('JPY')]);
  });

  it('allocates zero across every share', () => {
    expect(allocate(zero('JPY'), [1, 2])).toEqual([zero('JPY'), zero('JPY')]);
  });

  it('rejects a negative amount', () => {
    expect(() => allocate(money(-100, 'JPY'), [1])).toThrow(/negative amount/);
  });

  it('rejects an empty ratio list', () => {
    expect(() => allocate(money(100, 'JPY'), [])).toThrow(/must not be empty/);
  });

  it('rejects negative or non-finite ratios', () => {
    expect(() => allocate(money(100, 'JPY'), [1, -1])).toThrow(/non-negative/);
    expect(() => allocate(money(100, 'JPY'), [Number.NaN])).toThrow(/finite/);
  });

  it('rejects ratios that sum to zero', () => {
    expect(() => allocate(money(100, 'JPY'), [0, 0])).toThrow(/positive number/);
  });
});

describe('format', () => {
  it('formats a zero-decimal currency', () => {
    expect(format(money(1200, 'JPY'))).toContain('1,200');
  });

  it('shifts two-decimal currencies into major units', () => {
    expect(format(money(1234, 'USD'), 'en-US')).toBe('$12.34');
  });

  it('honours the requested locale', () => {
    expect(format(money(1000, 'EUR'), 'de-DE')).toContain('10,00');
  });
});
