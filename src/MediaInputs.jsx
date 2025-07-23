import { useContext } from "react";
import "./MediaInputs.css";
import DataContext from "./DataContext";

const MediaInputs = ({ info }) => {
  const { dispatch } = useContext(DataContext);
  return (
    <>
      <div className="CheckBoxGroup">
        {info.map(({ mediaType }, idx) => (
          <label key={mediaType}>
            <input
              id={idx}
              type="checkbox"
              onChange={() => dispatch({ type: "set-checks", idx })}
            />
            {mediaType}
          </label>
        ))}
      </div>
      <form id="MediaInputForm" className="MediaInputs">
        {info
          .filter((mediaType) => mediaType.show)
          .map(({ mediaType }) => (
            <label key={mediaType} htmlFor={mediaType}>
              {`${mediaType} Titles: `}
              <input
                id={mediaType}
                type="text"
                placeholder={`Input ${mediaType} Titles...`}
                onChange={(e) =>
                  dispatch({
                    type: "set-collect-text",
                    mediaType: mediaType,
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
