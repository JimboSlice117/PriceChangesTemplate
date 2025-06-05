const { getBStockInfo, conservativeNormalizeSku } = require('../PriceChangesTemplate');

describe('conservativeNormalizeSku', () => {
  test('basic normalization', () => {
    expect(conservativeNormalizeSku('abc-123')).toBe('ABC-123');
  });

  test('removes spaces and special characters', () => {
    expect(conservativeNormalizeSku('a$ b@c 123')).toBe('ABC123');
  });

  test('collapses multiple hyphens and trims', () => {
    expect(conservativeNormalizeSku('--abc--123--')).toBe('ABC-123');
  });

  test('handles numeric input', () => {
    expect(conservativeNormalizeSku(12345)).toBe('12345');
  });
});

describe('getBStockInfo', () => {
  test('detects BA suffix', () => {
    expect(getBStockInfo('ITEM-BA')).toEqual({ type: 'BA', multiplier: 0.95, sku: 'ITEM-BA', isSpecial: false });
  });

  test('detects prefix AA with special multiplier', () => {
    expect(getBStockInfo('AA-ITEM')).toEqual({ type: 'AA', multiplier: 0.98, sku: 'AA-ITEM', isSpecial: true });
  });

  test('handles NOACC and uses BC multiplier', () => {
    expect(getBStockInfo('ITEM-NOACC')).toEqual({ type: 'NOACC', multiplier: 0.85, sku: 'ITEM-NOACC', isSpecial: true });
  });

  test('returns null when no B-stock indicator', () => {
    expect(getBStockInfo('ITEM123')).toBeNull();
  });
});
