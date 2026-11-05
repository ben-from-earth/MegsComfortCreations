import CategoryTiles from '@/components/home/category-tiles';
import CustomOrnamentFeature from '@/components/home/custom-ornament-feature';
import HeroPaintRotator from '@/components/home/hero-paint-rotator';
import ShopStats from '@/components/home/shop-stats';

export default function Home() {
  return (
    <main className="min-h-screen w-full bg-[#f5eee2] font-(family-name:--font-nunito) text-[#1f1b18]">
      <HeroPaintRotator />
      <CategoryTiles />
      <CustomOrnamentFeature />
      <ShopStats />
    </main>
  );
}
