import { Fragment } from 'react';
import { homeStats } from '@/components/home/home-content';

export default function ShopStats() {
  return (
    <section
      aria-label="About the shop"
      className="mx-3 mt-3 flex items-center gap-16 bg-[#24493b] px-22.5 py-13 text-[#f5eee2]"
    >
      <div className="flex-[0_0_360px]">
        <p className="font-(family-name:--font-just-another-hand) text-5xl leading-[0.9] text-[#f6a27e]">
          {homeStats.title}
        </p>
        <p className="mt-3 text-[17px] leading-normal text-[#d9e2d6]">
          {homeStats.body}
        </p>
      </div>
      <span
        aria-hidden="true"
        className="w-px self-stretch bg-[rgba(245,238,226,0.25)]"
      />
      <dl className="m-0 flex flex-1 items-center justify-center gap-20">
        {homeStats.items.map((stat, index) => (
          <Fragment key={stat.label}>
            {index > 0 ? (
              <svg
                width="28"
                height="28"
                viewBox="0 0 24 24"
                aria-hidden="true"
              >
                <path
                  d="M12 0C13 8 16 11 24 12C16 13 13 16 12 24C11 16 8 13 0 12C8 11 11 8 12 0Z"
                  fill="#f6a27e"
                />
              </svg>
            ) : null}
            <div className="flex flex-col items-center">
              <dt className="mb-2 text-sm font-extrabold tracking-[3px] uppercase">
                {stat.label}
              </dt>
              <dd className="m-0 font-(family-name:--font-dm-serif-display) text-8xl leading-none">
                {stat.value}
              </dd>
            </div>
          </Fragment>
        ))}
      </dl>
    </section>
  );
}
