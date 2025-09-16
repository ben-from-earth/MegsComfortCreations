import { useDispatch } from "react-redux";
import { setChecks } from "@/state/collectorSlice";

const MediaCheckboxes = ({ mediaTypes, setSearchData }) => {
  //setup connection to redux slice
  const dispatch = useDispatch();

  return (
    <>
      <div className="m-6 flex flex-row content-center gap-5">
        {mediaTypes.map(({ type, label }, idx) => (
          <label
            key={type}
            className='font-["Just_Another_Hand"] text-[25px] tracking-wider'
          >
            <input
              className="m-[6px]"
              id={idx}
              type="checkbox"
              onChange={() => {
                dispatch(setChecks(idx));
                setSearchData((prev) => {
                  return prev.map((_, i) =>
                    i === idx ? { ...prev[i], text: "" } : prev[i],
                  );
                });
              }}
            />
            {`${label}s`}
          </label>
        ))}
      </div>
    </>
  );
};

export default MediaCheckboxes;
