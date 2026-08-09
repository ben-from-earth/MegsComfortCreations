import { requireAuthenticatedPageAccess } from 'lib/auth/page-access';
import ProfileClient from './ProfileClient';

export default async function ProfilePage() {
  const session = await requireAuthenticatedPageAccess();
  return <ProfileClient user={session.user} />;
}
