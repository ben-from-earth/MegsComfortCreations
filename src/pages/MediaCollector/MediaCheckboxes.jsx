import "./MediaCheckboxes.css";
import { useDispatch } from "react-redux";
import { setChecks } from "../../app/collectorSlice";

const MediaCheckboxes = ({ mediaTypes }) => {
  const dispatch = useDispatch();
  return (
    <>
      <div className="CheckBoxGroup">
        {mediaTypes.map(({ type, label }, idx) => (
          <label key={type} className="MCC-font">
            <input
              id={idx}
              type="checkbox"
              onChange={() => dispatch(setChecks(idx))}
            />
            {`${label}s`}
          </label>
        ))}
      </div>
    </>
  );
};

export default MediaCheckboxes;
