import { prefixedObjectKey, putPublicObject } from './local-image-storage';

export type OrderFileMimeType = 'image/png' | 'application/zip';

function safeOrderFileName(fileName: string) {
  const base = fileName.trim().split(/[\\/]/).pop() ?? '';
  const safe = base.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
  return safe || 'order-file';
}

export async function putOrderFile(input: {
  orderId: string;
  fileName: string;
  mimeType: OrderFileMimeType;
  body: Buffer;
}): Promise<{
  fileUrl: string;
  fileName: string;
  mimeType: OrderFileMimeType;
}> {
  const fileName = safeOrderFileName(input.fileName);
  const fileUrl = await putPublicObject({
    key: prefixedObjectKey('order-files', input.orderId, fileName),
    body: input.body,
    contentType: input.mimeType,
  });
  return { fileUrl, fileName, mimeType: input.mimeType };
}
