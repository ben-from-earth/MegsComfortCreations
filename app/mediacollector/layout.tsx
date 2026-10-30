import { Just_Another_Hand, Nunito } from 'next/font/google';
import { requireAdminPageAccess } from 'lib/auth/page-access';

const nunito = Nunito({ subsets: ['latin'] });
const justAnotherHand = Just_Another_Hand({
  weight: '400',
  subsets: ['latin'],
  variable: '--font-hand',
});

export default async function MediaCollectorLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  await requireAdminPageAccess();
  return (
    <div
      className={`${nunito.className} ${justAnotherHand.variable} min-h-screen bg-[#e5ebe0] text-[15px] text-[#2b2826]`}
    >
      {children}
    </div>
  );
}
