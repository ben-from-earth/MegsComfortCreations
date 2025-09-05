import "./QueryCounter.css";

const QueryCounter = () => {
  const qCount = Number(localStorage.getItem("queryCount"));
  return (
    <div className="QueryCounter">
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
