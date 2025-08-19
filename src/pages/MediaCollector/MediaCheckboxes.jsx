import "./MediaCheckboxes.css";
import { useDispatch } from "react-redux";
import { setChecks } from "../../state/collectorSlice";

const MediaCheckboxes = ({ mediaTypes, setSearchData }) => {
  //setup connection to redux slice
  const dispatch = useDispatch();

  return (
    <>
      <div className="CheckBoxGroup">
        {mediaTypes.map(({ type, label }, idx) => (
          <label key={type} className="MCC-font">
            <input
              id={idx}
              type="checkbox"
              onChange={() => {
                dispatch(setChecks(idx));
                setSearchData((prev) => {
                  return prev.map((_, i) =>
                    i === idx ? { ...prev[i], text: "" } : prev[i]
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
