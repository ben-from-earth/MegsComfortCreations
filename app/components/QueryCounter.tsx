"use client";

import { useEffect, useState } from "react";

const QueryCounter = () => {
  const [qCount, setqCount] = useState(0);
  useEffect(() => {
    setqCount(Number(localStorage.getItem("queryCount")));
  }, []);
  return (
    <div className="border-3 absolute right-0 top-0 m-4 rounded-lg border-darkpink bg-lightpink p-2.5 text-xl tracking-wider">
      <p>Query Count: {qCount}</p>
      {qCount > 100 ? (
        <p>Todays Cost: {`$${(0.005 * (qCount - 100)).toFixed(2)}`}</p>
      ) : (
        <></>
      )}
    </div>
  );
};

export default QueryCounter;
