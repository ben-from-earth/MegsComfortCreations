'use client';

import { useEffect, useState } from 'react';
import { trpc } from 'lib/trpc/client';

const QueryCounter = () => {
  const [qCount, setqCount] = useState(0);
  const date = new Date().toLocaleDateString('en-CA', {
    timeZone: 'America/New_York',
  });
  const { data: queryData } = trpc.database.getQueryCount.useQuery({ date });
  useEffect(() => {
    setqCount(queryData?.queryCount ?? 0);
  }, [queryData]);
  return (
    <div className="border-darkpink bg-lightpink absolute top-0 right-0 m-4 rounded-lg border-3 p-2.5 text-xl tracking-wider">
      <p>Today&apos;s Query Count: {qCount}</p>
      {qCount > 100 ? (
        <p>Today&apos;s Cost: {`$${(0.005 * (qCount - 100)).toFixed(2)}`}</p>
      ) : (
        <></>
      )}
    </div>
  );
};

export default QueryCounter;
