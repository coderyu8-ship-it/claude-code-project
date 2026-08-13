import { type Discount, totalDiscount } from './discount.js';
import {
  type Currency,
  type Money,
  CurrencyMismatchError,
  add,
  isNegative,
  money,
  multiply,
  subtract,
  zero,
} from './money.js';

export interface CartLine {
  readonly sku: string;
  readonly unitPrice: Money;
  readonly quantity: number;
}

export interface Cart {
  readonly currency: Currency;
  readonly lines: readonly CartLine[];
  readonly discounts?: readonly Discount[];
  /** Applied to the post-discount amount. Percent, not a fraction: 10 means 10%. */
  readonly taxRatePercent?: number;
}

export interface CartTotals {
  readonly subtotal: Money;
  readonly discountTotal: Money;
  readonly taxableBase: Money;
  readonly tax: Money;
  readonly total: Money;
}

function assertValidLine(line: CartLine, currency: Currency): void {
  if (line.unitPrice.currency !== currency) {
    throw new CurrencyMismatchError(currency, line.unitPrice.currency);
  }
  if (isNegative(line.unitPrice)) {
    throw new RangeError(`unit price of ${line.sku} must not be negative`);
  }
  if (!Number.isInteger(line.quantity) || line.quantity <= 0) {
    throw new RangeError(`quantity of ${line.sku} must be a positive integer, received ${line.quantity}`);
  }
}

export function lineTotal(line: CartLine): Money {
  return multiply(line.unitPrice, line.quantity);
}

export function subtotal(cart: Cart): Money {
  return cart.lines.reduce((sum, line) => {
    assertValidLine(line, cart.currency);
    return add(sum, lineTotal(line));
  }, zero(cart.currency));
}

export function calculateTotals(cart: Cart): CartTotals {
  const gross = subtotal(cart);
  const discountTotal = totalDiscount(gross, cart.discounts ?? []);
  const taxableBase = subtract(gross, discountTotal);

  const taxRatePercent = cart.taxRatePercent ?? 0;
  if (!Number.isFinite(taxRatePercent) || taxRatePercent < 0) {
    throw new RangeError(`taxRatePercent must be a non-negative number, received ${taxRatePercent}`);
  }

  const tax = multiply(taxableBase, taxRatePercent / 100);

  return {
    subtotal: gross,
    discountTotal,
    taxableBase,
    tax,
    total: add(taxableBase, tax),
  };
}

/** Line count, counting quantity — not the number of distinct SKUs. */
export function itemCount(cart: Cart): number {
  return cart.lines.reduce((count, line) => count + line.quantity, 0);
}

export function isEmpty(cart: Cart): boolean {
  return cart.lines.length === 0;
}

export function addLine(cart: Cart, line: CartLine): Cart {
  assertValidLine(line, cart.currency);

  const existing = cart.lines.find((candidate) => candidate.sku === line.sku);
  if (existing === undefined) {
    return { ...cart, lines: [...cart.lines, line] };
  }
  if (existing.unitPrice.amount !== line.unitPrice.amount) {
    throw new RangeError(`price conflict for ${line.sku}: cart holds a different unit price`);
  }

  return {
    ...cart,
    lines: cart.lines.map((candidate) =>
      candidate.sku === line.sku
        ? { ...candidate, quantity: candidate.quantity + line.quantity }
        : candidate,
    ),
  };
}

export function removeLine(cart: Cart, sku: string): Cart {
  return { ...cart, lines: cart.lines.filter((line) => line.sku !== sku) };
}

export function emptyCart(currency: Currency): Cart {
  return { currency, lines: [] };
}

/** Convenience for building a line without hand-writing a Money literal. */
export function line(sku: string, unitPriceMinorUnits: number, quantity: number, currency: Currency): CartLine {
  return { sku, unitPrice: money(unitPriceMinorUnits, currency), quantity };
}
