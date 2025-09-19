const QueryCounter = () => {
  const qCount = Number(localStorage.getItem("queryCount"));
  return (
    <div className='border-3 absolute right-0 top-0 m-4 rounded-[10px] border-[var(--darkpink)] bg-[var(--lightpink)] p-[10px] font-["Just_Another_Hand"] text-[20px] tracking-wider'>
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
