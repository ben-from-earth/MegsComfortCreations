const LEADING_ARTICLES = ['The', 'An', 'A'] as const;

function splitLeadingArticle(title: string) {
  const trimmed = title.trim();
  const lowerTitle = trimmed.toLowerCase();
  for (const article of LEADING_ARTICLES) {
    const lowerArticle = article.toLowerCase();
    if (lowerTitle.startsWith(`${lowerArticle} `)) {
      return {
        article: trimmed.slice(0, article.length),
        rest: trimmed.slice(article.length + 1).trim(),
      };
    }
    if (lowerTitle.endsWith(`, ${lowerArticle}`)) {
      return {
        article: trimmed.slice(-article.length),
        rest: trimmed.slice(0, -(article.length + 2)).trim(),
      };
    }
  }
  return null;
}

export function toStoredTitle(title: string): string {
  const split = splitLeadingArticle(title);
  return split ? `${split.rest}, ${split.article}` : title;
}

export function toDisplayTitle(title: string): string {
  const split = splitLeadingArticle(title);
  return split ? `${split.article} ${split.rest}` : title;
}

export function toTitleSearchTerm(title: string): string {
  return splitLeadingArticle(title)?.rest ?? title;
}
