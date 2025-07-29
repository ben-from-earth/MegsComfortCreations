import { useContext } from "react";
import "./MediaInputs.css";
import DataContext from "./DataContext";
import { TextField } from "@mui/material";

const MediaInputs = ({ mediaTypes }) => {
  const { dispatch } = useContext(DataContext);
  return (
    <>
      <form id="MediaInputForm" className="MediaInputs">
        {mediaTypes
          .filter((mediaType) => mediaType.show)
          .map(({ type }) => (
            <TextField
              className="MediaInput"
              id="outlined-multiline-static"
              multiline
              key={type}
              label={`${type} Titles`}
              rows={5}
              onChange={(e) =>
                dispatch({
                  type: "set-collect-text",
                  mediaType: type,
                  text: e.target.value,
                })
              }
            />
          ))}
      </form>
    </>
  );
};

export default MediaInputs;
