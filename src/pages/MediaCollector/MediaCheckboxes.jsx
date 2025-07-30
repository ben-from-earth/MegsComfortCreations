import { useContext } from "react";
import "./MediaCheckboxes.css";
import DataContext from "./DataContext";

const MediaCheckboxes = ({ mediaTypes }) => {
  const { dispatch } = useContext(DataContext);
  return (
    <>
      <div className="CheckBoxGroup">
        {mediaTypes.map(({ type }, idx) => (
          <label key={type} className="MCC-font">
            <input
              id={idx}
              type="checkbox"
              onChange={() => dispatch({ type: "set-checks", idx })}
            />
            {`${type}s`}
          </label>
        ))}
      </div>
    </>
  );
};

export default MediaCheckboxes;
