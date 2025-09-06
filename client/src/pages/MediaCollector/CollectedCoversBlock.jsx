import "./CollectedCoversBlock.css";
import BookIcon from "@mui/icons-material/BookTwoTone";
import MovieIcon from "@mui/icons-material/LocalMoviesTwoTone";
import VideoGameIcon from "@mui/icons-material/VideogameAssetTwoTone";
import AlbumIcon from "@mui/icons-material/AlbumTwoTone";
import IconButton from "@mui/material/IconButton";
import DeleteIcon from "@mui/icons-material/Delete";
import { useDispatch } from "react-redux";
import {
  addToDatabaseData,
  populateDatabaseData,
  removeFromDatabaseData,
  updateDatabaseData,
} from "../../state/databaseDataSlice";
import { memo, useContext, useEffect, useState } from "react";

//genre from context provider
import GenreContext from "../../context/GenreContext";
import {
  addToPNGCollectionList,
  removeFromPNGCollectionList,
} from "../../state/pngCollectionSlice";

// setup component Text area for each data field in the block
const MyTextArea = ({ name, label, type, blockID, value }) => {
  const dispatch = useDispatch();

  return (
    <>
      <label className="MCC-font" htmlFor={name}>
        {label}:
      </label>
      <textarea
        name={name}
        defaultValue={value}
        onChange={(e) => {
          dispatch(
            updateDatabaseData({
              blockID,
              type,
              name,
              newText: e.target.value,
            })
          );
        }}
      ></textarea>
    </>
  );
};

//setup memo so block doesnt rerender during other actions
const CollectedCoversBlock = memo(function CollectedCoversBlock({
  info: {
    type,
    images,
    blockInfo: {
      title,
      author,
      pub_year,
      page_count,
      spine_color = "#ffffff",
      databaseGenres = [],
    },
    blockID,
    isDatabase,
  },
  handleDeleteBlock,
}) {
  //setup connection to redux slice
  const dispatch = useDispatch();

  const icons = {
    book: <BookIcon className="Icon" />,
    movie: <MovieIcon className="Icon" />,
    video_game: <VideoGameIcon className="Icon" />,
    album: <AlbumIcon className="Icon" />,
  };
  const [color, setColor] = useState(spine_color);

  const bookSpecificPayload = {
    author,
    pub_year,
    page_count,
    genres: [],
  };

  const payload = {
    type,
    data: {
      title,
      spine_color: color,
      blockID,
      ...(type === "book" ? bookSpecificPayload : {}),
    },
  };

  //on mount, populate the database data (in the state) with the block information
  useEffect(() => {
    if (!isDatabase) {
      dispatch(populateDatabaseData(payload));
    } else {
      dispatch(addToPNGCollectionList({ type, spine_color, url: images[0] }));
    }

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dispatch]);

  //set up a local state to an array with an index for each image slot (current: 3) and set to false
  //this is for click tracking and the "selected" style and setting the image as a block datapoint
  const [clicked, setClicked] = useState(() =>
    Array(images.length).fill(false)
  );

  // function for handling an image click
  //this adds the image url to the database data (in the state) or removes it if its there already
  const handleClick = (blockID, type, idx, src) => {
    const next = !clicked[idx];
    setClicked((prev) =>
      prev.map((b, itemIndex) => (itemIndex === idx ? next : b))
    );
    if (next) {
      dispatch(
        addToDatabaseData({
          type,
          src,
          idx,
          blockID,
        })
      );
      dispatch(addToPNGCollectionList({ url: src, type, spine_color: color }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, idx }));
      dispatch(removeFromPNGCollectionList({ url: src }));
    }
  };

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  //if genre is clicked we add it to the data associated with the block and remove if unchecked
  const handleGenreClick = (genreText, checked) => {
    if (checked) {
      dispatch(addToDatabaseData({ blockID, type, genreText }));
    } else {
      dispatch(removeFromDatabaseData({ blockID, type, genreText }));
    }
  };

  //div under the covers to pick a color for the spine.
  //this is used in png creation
  const handleColorPick = async (blockID, type) => {
    if (!window.EyeDropper) {
      console.log("EyeDropper API not supported in this browser");
      return;
    }
    const eyeDropper = new EyeDropper();
    try {
      const { sRGBHex } = await eyeDropper.open();
      const spine_color = sRGBHex;
      setColor(spine_color);
      dispatch(addToDatabaseData({ blockID, type, spine_color }));
    } catch (e) {
      console.log(e);
    }
  };

  return (
    <div className={`Block ${type}`}>
      {isDatabase && <p className="databaseTag MCC-font">Database</p>}
      {icons[type]}
      <IconButton
        aria-label="delete"
        className="DeleteIcon"
        onClick={() =>
          handleDeleteBlock({ blockID, type, deleteBlock: true, urls: images })
        }
      >
        <DeleteIcon />
      </IconButton>
      <div className="imageContainer">
        {images.map((src, idx) => (
          <div
            className="imageWrapper"
            key={src}
            onClick={() => {
              if (!isDatabase) {
                handleClick(blockID, type, idx, src);
              }
            }}
          >
            <img className={`${type}-img `} src={src}></img>
            <div className={`overlay ${clicked[idx] ? "show" : ""}`}>
              <p>Selected</p>
            </div>
          </div>
        ))}
      </div>
      {type !== "album" ? (
        <div
          className="colorPicker"
          style={{ backgroundColor: color }}
          onClick={() => handleColorPick(blockID, type)}
        ></div>
      ) : (
        <></>
      )}
      <div className="titleInfoContainer">
        <MyTextArea
          name="title"
          label="Title"
          type={type}
          dispatch={dispatch}
          blockID={blockID}
          value={title || ""}
        />
        {type === "book" ? (
          <>
            <MyTextArea
              name="author"
              label="Author"
              type={type}
              blockID={blockID}
              value={author || ""}
            />
            <MyTextArea
              name="pub_year"
              label="Publication Year"
              type={type}
              blockID={blockID}
              value={pub_year || ""}
            />
            <MyTextArea
              name="page_count"
              label="Page Count"
              type={type}
              blockID={blockID}
              value={page_count || ""}
            />
          </>
        ) : (
          <></>
        )}
      </div>
      {type === "book" ? (
        <div className="genreCheckboxes">
          {genres?.map((text, idx) => (
            <label
              key={idx}
              name={text}
              className="MCC-font"
              onChange={(e) => {
                handleGenreClick(text, e.target.checked);
              }}
            >
              <input
                type="checkbox"
                defaultChecked={databaseGenres?.includes(text)}
              />
              {text}
            </label>
          ))}
        </div>
      ) : (
        <></>
      )}
    </div>
  );
});

export default CollectedCoversBlock;
