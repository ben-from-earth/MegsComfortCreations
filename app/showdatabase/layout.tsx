import { requireAdminPageAccess } from 'lib/auth/page-access';

export default async function ShowDatabaseLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  await requireAdminPageAccess();
  return children;
}
