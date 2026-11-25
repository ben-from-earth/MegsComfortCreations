import { putPublicObject } from './local-image-storage';

const ORDER_FILE_MIME_TYPES = ['image/png', 'application/zip'] as const;

export type OrderFileMimeType = (typeof ORDER_FILE_MIME_TYPES)[number];

function isOrderFileMimeType(mimeType: string): mimeType is OrderFileMimeType {
  return ORDER_FILE_MIME_TYPES.some((allowed) => allowed === mimeType);
}

function safeOrderFileName(fileName: string) {
  const base = fileName.trim().split(/[\\/]/).pop() ?? '';
  const safe = base.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
  return safe || 'order-file';
}

export async function putOrderFile(input: {
  orderId: string;
  fileName: string;
  mimeType: string;
  body: Buffer;
}): Promise<{ fileUrl: string; fileName: string; mimeType: OrderFileMimeType }> {
  if (!isOrderFileMimeType(input.mimeType)) {
    throw new Error(
      `Order file type is not allowed. Received "${input.mimeType}".`,
    );
  }
  if (input.body.length === 0) {
    throw new Error('Order file is empty.');
  }

  const fileName = safeOrderFileName(input.fileName);
  const key = `order-files/${input.orderId}/${fileName}`;
  const fileUrl = await putPublicObject({
    key,
    body: input.body,
    contentType: input.mimeType,
  });
  return { fileUrl, fileName, mimeType: input.mimeType };
}
