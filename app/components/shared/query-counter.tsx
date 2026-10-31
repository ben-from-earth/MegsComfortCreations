'use client';

import { trpc } from 'lib/trpc/client';

const FREE_SEARCHES_PER_DAY = 100;
const NEAR_LIMIT_SEARCH_COUNT = 90;
const COST_PER_PAID_SEARCH = 0.005;

const QueryCounter = () => {
  const date = new Date().toLocaleDateString('en-CA', {
    timeZone: 'America/New_York',
  });
  const { data: queryData } = trpc.database.getQueryCount.useQuery({ date });
  const queryCount = queryData?.queryCount ?? 0;
  const paidSearchCount = Math.max(queryCount - FREE_SEARCHES_PER_DAY, 0);
  const isOverLimit = paidSearchCount > 0;
  const barColorClass = isOverLimit
    ? 'bg-[#b0533a]'
    : queryCount >= NEAR_LIMIT_SEARCH_COUNT
      ? 'bg-[#b8862f]'
      : 'bg-[#6f7f66]';
  const barWidthPercent =
    (Math.min(queryCount, FREE_SEARCHES_PER_DAY) / FREE_SEARCHES_PER_DAY) * 100;

  return (
    <div
      className={`flex items-center gap-3 rounded-xl border bg-white px-3.5 py-2 text-[#2b2826] ${
        isOverLimit ? 'border-[#e2b8a8]' : 'border-[#e6e0d8]'
      }`}
    >
      <div className="flex min-w-[150px] flex-col gap-1.25">
        <span
          className={`text-xs font-extrabold ${
            isOverLimit ? 'text-[#9a3b26]' : 'text-[#6b645d]'
          }`}
        >
          {isOverLimit
            ? `${paidSearchCount} paid ${paidSearchCount === 1 ? 'search' : 'searches'}`
            : 'Free searches today'}
        </span>
        <div className="h-1.5 overflow-hidden rounded-full bg-[#ece6de]">
          <div
            className={`h-full ${barColorClass}`}
            style={{ width: `${barWidthPercent}%` }}
          />
        </div>
      </div>
      <div className="flex flex-col items-end">
        <span className="text-lg font-extrabold">
          {queryCount}
          {isOverLimit ? null : (
            <span className="text-[13px] font-bold text-[#6b645d]">
              {' '}
              / {FREE_SEARCHES_PER_DAY}
            </span>
          )}
        </span>
        {isOverLimit ? (
          <span className="text-xs font-bold text-[#9a3b26]">
            ${(COST_PER_PAID_SEARCH * paidSearchCount).toFixed(2)} today
          </span>
        ) : null}
      </div>
    </div>
  );
};

export default QueryCounter;
