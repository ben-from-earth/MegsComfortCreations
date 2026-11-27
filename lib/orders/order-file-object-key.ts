function decodedPathSegments(pathname: string): string[] | null {
  const segments: string[] = [];
  for (const segment of pathname.split('/')) {
    if (segment.length === 0) {
      continue;
    }
    try {
      segments.push(decodeURIComponent(segment));
    } catch {
      return null;
    }
  }
  return segments;
}

export function orderFileObjectKey(
  fileUrl: string,
  publicBaseUrl: string,
): string | null {
  let url: URL;
  let base: URL;
  try {
    url = new URL(fileUrl);
    base = new URL(publicBaseUrl);
  } catch {
    return null;
  }
  if (url.origin !== base.origin) {
    return null;
  }

  const baseSegments = decodedPathSegments(base.pathname);
  const fileSegments = decodedPathSegments(url.pathname);
  if (!baseSegments || !fileSegments) {
    return null;
  }
  if (fileSegments.length <= baseSegments.length) {
    return null;
  }
  for (let index = 0; index < baseSegments.length; index += 1) {
    if (fileSegments[index] !== baseSegments[index]) {
      return null;
    }
  }

  const keySegments = fileSegments.slice(baseSegments.length);
  if (!keySegments.includes('order-files')) {
    return null;
  }
  return keySegments.join('/');
}
