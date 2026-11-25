export type StoredOrderIdentity = {
  id: string;
  orderNumber: string;
};

export type OrderWriteChoice =
  | { action: 'insert' }
  | { action: 'update'; orderId: string }
  | { action: 'conflict' }
  | { action: 'missing' };

export function chooseOrderWrite(input: {
  requestedOrderId: string | null;
  orderWithRequestedId: StoredOrderIdentity | null;
  orderWithRequestedNumber: StoredOrderIdentity | null;
}): OrderWriteChoice {
  if (input.requestedOrderId !== null && input.orderWithRequestedId === null) {
    return { action: 'missing' };
  }

  const numberOwner = input.orderWithRequestedNumber;
  if (
    numberOwner &&
    input.requestedOrderId !== null &&
    numberOwner.id !== input.requestedOrderId
  ) {
    return { action: 'conflict' };
  }

  if (input.requestedOrderId !== null) {
    return { action: 'update', orderId: input.requestedOrderId };
  }

  if (numberOwner) {
    return { action: 'update', orderId: numberOwner.id };
  }

  return { action: 'insert' };
}
