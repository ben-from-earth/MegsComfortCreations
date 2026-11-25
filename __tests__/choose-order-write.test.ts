import { chooseOrderWrite } from 'lib/orders/choose-order-write';

const existing = { id: 'order-a', orderNumber: '1001' };
const other = { id: 'order-b', orderNumber: '1002' };

describe('chooseOrderWrite', () => {
  test.each([
    ['inserts a new order number', null, null, null, { action: 'insert' }],
    [
      'updates the row that already has the order number',
      null,
      null,
      existing,
      { action: 'update', orderId: 'order-a' },
    ],
    [
      'updates the requested order when its number changes',
      'order-a',
      existing,
      null,
      { action: 'update', orderId: 'order-a' },
    ],
    [
      'conflicts when the number belongs to a different order',
      'order-a',
      existing,
      other,
      { action: 'conflict' },
    ],
    [
      'reports a requested order that is not stored',
      'order-a',
      null,
      other,
      { action: 'missing' },
    ],
  ])(
    '%s',
    (
      _,
      requestedOrderId,
      orderWithRequestedId,
      orderWithRequestedNumber,
      expected,
    ) => {
      expect(
        chooseOrderWrite({
          requestedOrderId,
          orderWithRequestedId,
          orderWithRequestedNumber,
        }),
      ).toEqual(expected);
    },
  );
});
