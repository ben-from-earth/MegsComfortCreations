import axios from 'axios';
import { createHash } from 'node:crypto';
import { PutObjectCommand, S3Client, S3ClientConfig } from '@aws-sdk/client-s3';
import { MediaType } from 'lib/constants/mediaTypes';
import { isExternalImageUrl, normalizeImagePath } from './image-path-utils';

const DEFAULT_MAX_IMAGE_BYTES = 5 * 1024 * 1024;
const DEFAULT_DOWNLOAD_TIMEOUT_MS = 15000;
const DEFAULT_UPLOAD_NAMESPACE = 'shared';
const DEFAULT_S3_KEY_PREFIX = 'media-images';
const DEFAULT_ALLOWED_MIME_TYPES = [
  'image/jpeg',
  'image/png',
  'image/webp',
  'image/gif',
];

const MIME_EXTENSION_MAP: Record<string, string> = {
  'image/jpeg': 'jpg',
  'image/png': 'png',
  'image/webp': 'webp',
  'image/gif': 'gif',
};

export type PersistedImageFile = {
  publicPath: string;
  mimeType: string | null;
  sizeBytes: number | null;
  sourceUrl: string;
};

export type PersistExternalImageInput = {
  sourceUrl: string;
  mediaType: MediaType;
  mediaId: string;
  sortOrder: number;
};

type S3StorageConfig = {
  client: S3Client;
  bucket: string;
  publicBaseUrl: string;
  keyPrefix: string;
};

let cachedS3Client: S3Client | null = null;
let cachedS3ClientCacheKey: string | null = null;

function getAllowedMimeTypes(): string[] {
  const fromEnv = process.env.ALLOWED_IMAGE_MIME_TYPES
    ?.split(',')
    .map((value) => value.trim())
    .filter(Boolean);
  if (!fromEnv || fromEnv.length === 0) {
    return DEFAULT_ALLOWED_MIME_TYPES;
  }
  return fromEnv;
}

function getMaxImageBytes(): number {
  const parsed = Number(process.env.MAX_IMAGE_BYTES);
  return Number.isFinite(parsed) && parsed > 0
    ? parsed
    : DEFAULT_MAX_IMAGE_BYTES;
}

function getDownloadTimeoutMs(): number {
  const parsed = Number(process.env.IMAGE_DOWNLOAD_TIMEOUT_MS);
  return Number.isFinite(parsed) && parsed > 0
    ? parsed
    : DEFAULT_DOWNLOAD_TIMEOUT_MS;
}

function getUploadNamespace(): string {
  const explicitNamespace = process.env.IMAGE_UPLOAD_NAMESPACE?.trim();
  if (explicitNamespace) {
    return explicitNamespace;
  }
  const nodeEnvironment = process.env.NODE_ENV?.trim();
  if (nodeEnvironment) {
    return nodeEnvironment;
  }
  return DEFAULT_UPLOAD_NAMESPACE;
}

function parseBooleanEnv(value: string | undefined): boolean {
  if (!value) {
    return false;
  }
  const normalizedValue = value.trim().toLowerCase();
  return normalizedValue === '1' || normalizedValue === 'true';
}

function normalizeBaseUrl(url: string): string {
  return url.replace(/\/+$/, '');
}

function normalizePathSegment(segment: string | undefined): string {
  if (!segment) {
    return '';
  }
  return segment.trim().replace(/^\/+|\/+$/g, '');
}

function getRequiredEnvValue(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable "${name}" for S3 image storage.`);
  }
  return value;
}

function createS3Client(): S3Client {
  const region = getRequiredEnvValue('S3_REGION');
  const endpoint = process.env.S3_ENDPOINT?.trim();
  const forcePathStyle = parseBooleanEnv(process.env.S3_FORCE_PATH_STYLE);
  const accessKeyId = process.env.AWS_ACCESS_KEY_ID?.trim();
  const secretAccessKey = process.env.AWS_SECRET_ACCESS_KEY?.trim();
  const sessionToken = process.env.AWS_SESSION_TOKEN?.trim();

  const cacheKey = JSON.stringify({
    region,
    endpoint: endpoint ?? '',
    forcePathStyle,
    accessKeyId: accessKeyId ?? '',
    hasSecretAccessKey: Boolean(secretAccessKey),
    hasSessionToken: Boolean(sessionToken),
  });
  if (cachedS3Client && cachedS3ClientCacheKey === cacheKey) {
    return cachedS3Client;
  }

  const config: S3ClientConfig = {
    region,
    forcePathStyle,
  };
  if (endpoint) {
    config.endpoint = endpoint;
  }
  if (accessKeyId && secretAccessKey) {
    config.credentials = {
      accessKeyId,
      secretAccessKey,
      ...(sessionToken ? { sessionToken } : {}),
    };
  }

  cachedS3Client = new S3Client(config);
  cachedS3ClientCacheKey = cacheKey;
  return cachedS3Client;
}

function getS3StorageConfig(): S3StorageConfig {
  const bucket = getRequiredEnvValue('S3_BUCKET');
  const publicBaseUrl = normalizeBaseUrl(getRequiredEnvValue('S3_PUBLIC_BASE_URL'));
  const keyPrefix =
    normalizePathSegment(process.env.S3_KEY_PREFIX) || DEFAULT_S3_KEY_PREFIX;
  return {
    client: createS3Client(),
    bucket,
    publicBaseUrl,
    keyPrefix,
  };
}

function resolveMimeType(contentTypeHeader: string | undefined): string | null {
  if (!contentTypeHeader) {
    return null;
  }
  return contentTypeHeader.split(';')[0]?.trim().toLowerCase() ?? null;
}

function resolveExtension(mimeType: string | null): string | null {
  if (!mimeType) {
    return null;
  }
  return MIME_EXTENSION_MAP[mimeType] ?? null;
}

type FileTypeDetection = { ext: string; mime: string } | null;

async function detectFileTypeFromBuffer(imageBuffer: Buffer): Promise<FileTypeDetection> {
  try {
    const { fileTypeFromBuffer } = await import('file-type');
    return (await fileTypeFromBuffer(imageBuffer)) ?? null;
  } catch (error) {
    // Keep router/module loading resilient in test environments where the optional
    // detector package may be unavailable, and fall back to response headers.
    if (error instanceof Error && error.message.includes("Cannot find module 'file-type'")) {
      return null;
    }
    throw error;
  }
}

export async function persistExternalImageToS3(
  input: PersistExternalImageInput,
): Promise<PersistedImageFile> {
  const sourceUrl = normalizeImagePath(input.sourceUrl);
  if (!isExternalImageUrl(sourceUrl)) {
    return {
      publicPath: sourceUrl,
      mimeType: null,
      sizeBytes: null,
      sourceUrl,
    };
  }

  const maxImageBytes = getMaxImageBytes();
  const response = await axios.get<ArrayBuffer>(sourceUrl, {
    responseType: 'arraybuffer',
    timeout: getDownloadTimeoutMs(),
    validateStatus: (status) => status >= 200 && status < 300,
    maxContentLength: maxImageBytes,
  });

  const imageBuffer = Buffer.from(response.data);
  if (imageBuffer.length > maxImageBytes) {
    throw new Error(
      `Image download exceeded MAX_IMAGE_BYTES for "${sourceUrl}" (${imageBuffer.length} bytes).`,
    );
  }

  const detectedFileType = await detectFileTypeFromBuffer(imageBuffer);
  const headerMimeType = resolveMimeType(response.headers['content-type']);
  const mimeType = detectedFileType?.mime ?? headerMimeType ?? null;
  const allowedMimeTypes = getAllowedMimeTypes();
  if (!mimeType || !allowedMimeTypes.includes(mimeType)) {
    throw new Error(
      `Image MIME type is not allowed for "${sourceUrl}". Received "${mimeType ?? 'unknown'}".`,
    );
  }

  const extension = detectedFileType?.ext ?? resolveExtension(mimeType);
  if (!extension) {
    throw new Error(`Unable to resolve file extension for "${sourceUrl}".`);
  }

  const now = new Date();
  const year = String(now.getUTCFullYear());
  const month = String(now.getUTCMonth() + 1).padStart(2, '0');
  const uploadNamespace = getUploadNamespace();

  const fileHash = createHash('sha256')
    .update(imageBuffer)
    .digest('hex')
    .slice(0, 16);
  const orderToken = String(input.sortOrder).padStart(2, '0');
  const fileName = `${input.mediaType}-${input.mediaId}-${orderToken}-${fileHash}.${extension}`;
  const s3Config = getS3StorageConfig();
  const objectKey = [
    s3Config.keyPrefix,
    uploadNamespace,
    'covers',
    year,
    month,
    fileName,
  ]
    .filter(Boolean)
    .join('/');

  await s3Config.client.send(
    new PutObjectCommand({
      Bucket: s3Config.bucket,
      Key: objectKey,
      Body: imageBuffer,
      ContentType: mimeType,
      CacheControl: 'public, max-age=31536000, immutable',
    }),
  );

  return {
    publicPath: `${s3Config.publicBaseUrl}/${objectKey}`,
    mimeType,
    sizeBytes: imageBuffer.length,
    sourceUrl,
  };
}

// Temporary compatibility export to avoid broad call-site churn.
export const persistExternalImageToLocalDisk = persistExternalImageToS3;
