import { Just_Another_Hand, Nunito } from 'next/font/google';

const nunito = Nunito({ subsets: ['latin'] });
const justAnotherHand = Just_Another_Hand({
  weight: '400',
  subsets: ['latin'],
  variable: '--font-hand',
});

export default function ComingSoon({ title }: { title: string }) {
  return (
    <div
      className={`${nunito.className} ${justAnotherHand.variable} relative isolate min-h-screen bg-[#e5ebe0] text-[15px] text-[#2b2826]`}
    >
      <div
        aria-hidden
        className="fixed inset-0 -z-10 bg-[url('/collector-background.png')] bg-cover bg-center"
      />
      <main className="flex justify-center px-6 py-20">
        <section className="w-full max-w-xl overflow-hidden rounded-[20px] border border-[#e6e0d8] bg-white text-center shadow-[0_8px_30px_rgba(60,50,40,0.07)]">
          <div className="border-b border-[#e3e9de] bg-[#f3f6f0] px-7 py-4.5">
            <h1 className="m-0 font-(family-name:--font-hand) text-[52px] leading-[0.9] font-normal tracking-[1px]">
              {title} coming soon
            </h1>
          </div>
          <p className="m-0 px-7 py-8 text-lg font-bold">Stay Tuned!</p>
        </section>
      </main>
    </div>
  );
}
