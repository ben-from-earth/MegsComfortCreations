import { TextField } from '@mui/material';

//This component is mainly the data collection of the app.
//Based on which media checkboxes are checked, the text area will show.
//Inputs are a comma seperated list of the wanted titles (books require title / author list)
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
              className="w-90 rounded-sm bg-white"
              id={`outlined-multiline-static ${label}`}
              multiline
              key={type}
              label={`${label} Titles`}
              slotProps={{
                inputLabel: {
                  sx: {
                    '&.MuiInputLabel-shrink': {
                      backgroundColor: 'white',
                      borderRadius: '8px',
                      px: '10px',
                      color: 'rgb(0,0,0, 0.5)',
                      transform: 'translate(6px, -8px) scale(0.75)',
                    },
                  },
                },
              }}
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
