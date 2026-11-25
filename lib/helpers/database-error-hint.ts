export function databaseErrorHint(error: unknown) {
  const parts: string[] = [];
  const visit = (value: unknown, depth: number) => {
    if (depth > 4 || value === null || value === undefined) {
      return;
    }
    if (value instanceof Error) {
      parts.push(value.message);
      visit(value.cause, depth + 1);
    }
    if (typeof value === 'object' && value !== null && 'constraint' in value) {
      const constraint = value.constraint;
      if (typeof constraint === 'string') {
        parts.push(constraint);
      }
    }
  };
  visit(error, 0);
  return parts.join(' ');
}
