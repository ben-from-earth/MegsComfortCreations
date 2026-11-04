import { notFound } from 'next/navigation';
import ComingSoon from '@/components/shared/coming-soon';
import { SHOP_DIVISIONS } from '@/shop/shop-divisions';

export const dynamicParams = false;

export function generateStaticParams() {
  return SHOP_DIVISIONS.map(({ slug }) => ({ division: slug }));
}

export default async function ShopDivisionPage({
  params,
}: {
  params: Promise<{ division: string }>;
}) {
  const { division } = await params;
  const shopDivision = SHOP_DIVISIONS.find(({ slug }) => slug === division);
  if (!shopDivision) notFound();

  return <ComingSoon title={shopDivision.label} />;
}
