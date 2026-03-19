const HTTP_URL_PATTERN = /^https?:\/\//i;

function normalizeBaseUrl(url: string): string {
  return url.replace(/\/+$/, '');
}

function getManagedImageBaseUrls(): string[] {
  const configuredPublicBaseUrl = process.env.S3_PUBLIC_BASE_URL?.trim();
  if (!configuredPublicBaseUrl) {
    return [];
  }
  return [normalizeBaseUrl(configuredPublicBaseUrl)];
}

function isManagedImageUrl(imageUrl: string): boolean {
  const normalizedImageUrl = imageUrl.trim();
  if (!HTTP_URL_PATTERN.test(normalizedImageUrl)) {
    return false;
  }
  return getManagedImageBaseUrls().some((baseUrl) =>
    normalizedImageUrl.startsWith(`${baseUrl}/`),
  );
}

export function isExternalImageUrl(imageUrl: string): boolean {
  const normalizedImageUrl = imageUrl.trim();
  return (
    HTTP_URL_PATTERN.test(normalizedImageUrl) &&
    !isManagedImageUrl(normalizedImageUrl)
  );
}

export function isLocalImagePath(imageUrl: string): boolean {
  const normalizedImageUrl = imageUrl.trim();
  return (
    normalizedImageUrl.startsWith('/uploads/') ||
    isManagedImageUrl(normalizedImageUrl)
  );
}

export function normalizeImagePath(imageUrl: unknown): string {
  if (typeof imageUrl === 'string') {
    return imageUrl.trim();
  }
  if (imageUrl && typeof imageUrl === 'object') {
    const record = imageUrl as { url?: unknown; src?: unknown };
    if (typeof record.url === 'string') {
      return record.url.trim();
    }
    if (typeof record.src === 'string') {
      return record.src.trim();
    }
  }
  return '';
}
