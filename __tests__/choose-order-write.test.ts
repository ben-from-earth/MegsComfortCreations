import { chooseOrderWrite } from 'lib/orders/choose-order-write';

const existing = { id: 'order-a', orderNumber: '1001' };
const other = { id: 'order-b', orderNumber: '1002' };

describe('chooseOrderWrite', () => {
  test('inserts when the order number is new', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: null,
        orderWithRequestedId: null,
        orderWithRequestedNumber: null,
      }),
    ).toEqual({ action: 'insert' });
  });

  test('updates the row that already has the order number', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: null,
        orderWithRequestedId: null,
        orderWithRequestedNumber: existing,
      }),
    ).toEqual({ action: 'update', orderId: 'order-a' });
  });

  test('updates the requested order when its number changes', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: 'order-a',
        orderWithRequestedId: existing,
        orderWithRequestedNumber: null,
      }),
    ).toEqual({ action: 'update', orderId: 'order-a' });
  });

  test('updates the requested order when the number stays on that row', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: 'order-a',
        orderWithRequestedId: existing,
        orderWithRequestedNumber: existing,
      }),
    ).toEqual({ action: 'update', orderId: 'order-a' });
  });

  test('conflicts when the number belongs to a different order', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: 'order-a',
        orderWithRequestedId: existing,
        orderWithRequestedNumber: other,
      }),
    ).toEqual({ action: 'conflict' });
  });

  test('reports a requested order that is not stored', () => {
    expect(
      chooseOrderWrite({
        requestedOrderId: 'order-a',
        orderWithRequestedId: null,
        orderWithRequestedNumber: other,
      }),
    ).toEqual({ action: 'missing' });
  });
});
