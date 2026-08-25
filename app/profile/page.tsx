import { requireAuthenticatedPageAccess } from 'lib/auth/page-access';
import ProfileClient from './profile-client';

export default async function ProfilePage() {
  const session = await requireAuthenticatedPageAccess();
  return <ProfileClient user={session.user} />;
}
