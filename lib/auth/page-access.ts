import { auth } from 'lib/auth';
import {
  decideAdminPageAccess,
  decideAuthenticatedPageAccess,
} from 'lib/auth/page-access-decisions';
import { headers } from 'next/headers';
import { redirect } from 'next/navigation';

export async function requireAdminPageAccess() {
  const session = await auth.api.getSession({ headers: await headers() });
  const redirectTo = decideAdminPageAccess(session);
  if (redirectTo) redirect(redirectTo);
  return session!;
}

export async function requireAuthenticatedPageAccess() {
  const session = await auth.api.getSession({ headers: await headers() });
  const redirectTo = decideAuthenticatedPageAccess(session);
  if (redirectTo) redirect(redirectTo);
  return session!;
}
