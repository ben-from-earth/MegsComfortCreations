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

export function normalizeImagePath(imageUrl: string): string {
  return imageUrl.trim();
}
