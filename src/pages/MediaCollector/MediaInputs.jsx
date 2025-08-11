import "./MediaInputs.css";
import { TextField } from "@mui/material";

const MediaInputs = ({ mediaTypes, setSearchData }) => {
  return (
    <>
      <form id="MediaInputForm" className="MediaInputs">
        {mediaTypes
          .filter((mediaType) => mediaType.show)
          .map(({ type, label }) => (
            <TextField
              className="MediaInput"
              id="outlined-multiline-static"
              multiline
              key={type}
              label={`${label} Titles`}
              rows={5}
              onChange={(e) => {
                setSearchData((prev) => {
                  return prev.map((media) =>
                    media.type === type
                      ? { type: type, text: e.target.value }
                      : media
                  );
                });
              }}
            />
          ))}
      </form>
    </>
  );
};

export default MediaInputs;
