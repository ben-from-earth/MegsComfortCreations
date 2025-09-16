import { TextField } from "@mui/material";

const MediaInputs = ({ mediaTypes, setSearchData }) => {
  return (
    <>
      <form
        id="MediaInputForm"
        className="MediaInputs flex flex-col items-center gap-4 p-5 sm:grid sm:grid-cols-2"
      >
        {mediaTypes
          .filter((mediaType) => mediaType.show)
          .map(({ type, label }) => (
            <TextField
              className="w-75 rounded-sm bg-white"
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
                      : media,
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
