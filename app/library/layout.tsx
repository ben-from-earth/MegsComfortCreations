import { requireAdminPageAccess } from 'lib/auth/page-access';

export default async function LibraryLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  await requireAdminPageAccess();
  return children;
}
