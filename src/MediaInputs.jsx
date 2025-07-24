import { useContext } from "react";
import "./MediaInputs.css";
import DataContext from "./DataContext";

const MediaInputs = ({ info: { mediaTypes } }) => {
  const { dispatch } = useContext(DataContext);
  return (
    <>
      <div className="CheckBoxGroup">
        {mediaTypes.map(({ type }, idx) => (
          <label key={type}>
            <input
              id={idx}
              type="checkbox"
              onChange={() => dispatch({ type: "set-checks", idx })}
            />
            {type}
          </label>
        ))}
      </div>
      <form id="MediaInputForm" className="MediaInputs">
        {mediaTypes
          .filter((mediaType) => mediaType.show)
          .map(({ type }) => (
            <label key={type} htmlFor={type}>
              {`${type} Titles: `}
              <input
                id={type}
                type="text"
                placeholder={`Input ${type} Titles...`}
                onChange={(e) =>
                  dispatch({
                    type: "set-collect-text",
                    mediaType: type,
                    text: e.target.value,
                  })
                }
              />
            </label>
          ))}
      </form>
    </>
  );
};

export default MediaInputs;
