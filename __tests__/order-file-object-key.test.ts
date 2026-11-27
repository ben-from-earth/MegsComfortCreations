import { orderFileObjectKey } from 'lib/orders/order-file-object-key';

const base = 'https://cdn.example.com';
const orderFile = `${base}/media-images/shared/order-files/order-id/sheet.png`;

describe('orderFileObjectKey', () => {
  test('returns the key for an order file on our host', () => {
    expect(orderFileObjectKey(orderFile, base)).toBe(
      'media-images/shared/order-files/order-id/sheet.png',
    );
  });

  test('drops a base path that belongs to the public url', () => {
    expect(
      orderFileObjectKey(
        'https://cdn.example.com/files/media-images/shared/order-files/order-id/sheet.png',
        'https://cdn.example.com/files',
      ),
    ).toBe('media-images/shared/order-files/order-id/sheet.png');
  });

  test('ignores a cover on the same host', () => {
    expect(
      orderFileObjectKey(
        `${base}/media-images/shared/covers/2026/01/book.jpg`,
        base,
      ),
    ).toBeNull();
  });

  test('ignores an external url', () => {
    expect(
      orderFileObjectKey(
        'https://images.example.org/order-files/sheet.png',
        base,
      ),
    ).toBeNull();
  });

  test('ignores a host that only shares a prefix', () => {
    expect(
      orderFileObjectKey(
        'https://cdn.example.com.evil/media-images/shared/order-files/order-id/sheet.png',
        base,
      ),
    ).toBeNull();
  });

  test('ignores a path that only shares a prefix', () => {
    expect(
      orderFileObjectKey(
        `${base}/media-images/shared/order-files/order-id/sheet.png`,
        `${base}/media`,
      ),
    ).toBeNull();
  });

  test('ignores a path that climbs out of the order file prefix', () => {
    expect(
      orderFileObjectKey(
        `${base}/media-images/shared/order-files/../covers/secret.png`,
        base,
      ),
    ).toBeNull();
  });
});
