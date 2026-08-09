export type PageSessionUser = {
  role?: string | null;
};
export type PageSession = {
  user: PageSessionUser;
};
export type PageAccessRedirect = '/login' | '/';

export function decideAdminPageAccess(
  session: PageSession | null,
): PageAccessRedirect | null {
  if (!session) return '/login';
  if (session.user.role !== 'admin') return '/';
  return null;
}

export function decideAuthenticatedPageAccess(
  session: PageSession | null,
): Extract<PageAccessRedirect, '/login'> | null {
  if (!session) return '/login';
  return null;
}
