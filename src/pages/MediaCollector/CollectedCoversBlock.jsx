import "./CollectedCoversBlock.css";
import BookIcon from "@mui/icons-material/BookTwoTone";
import MovieIcon from "@mui/icons-material/LocalMoviesTwoTone";
import VideoGameIcon from "@mui/icons-material/VideogameAssetTwoTone";
import AlbumIcon from "@mui/icons-material/AlbumTwoTone";
import { useDispatch, useSelector } from "react-redux";
import {
  addImageToDatabaseData,
  populateDatabaseData,
  removeImageFromDatabaseData,
  selectDatabaseData,
  updateDatabaseData,
} from "../../state/databaseDataSlice";
import { memo, useContext, useEffect, useState } from "react";

//genre from context provider
import GenreContext from "../../context/GenreContext";

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  //setup memo so block doesnt rerender during other actions
  info: {
    type,
    images,
    blockInfo: { title, author, pub_year, page_count },
    blockID,
  },
}) {
  // setup component Text area for each data field in the block
  const MyTextArea = ({ name, label }) => {
    //setup connection to redux slice
    const dispatch = useDispatch();
    const databaseData = useSelector(selectDatabaseData);
    // databaseData: [{ type, label, data: [...] }]

    const typeData = databaseData.find((media) => media.type === type);
    //typeData: [{ title, blockID, images, ... }]

    const block = typeData?.data?.find((data) => data.blockID === blockID);
    // Use store value if present; otherwise fallback from props
    const value = block ? block[name] : "";

    return (
      <>
        <label className="MCC-font" htmlFor={name}>
          {label}:
        </label>
        <textarea
          name={name}
          value={value}
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
        />
      </>
    );
  };

  const icons = {
    book: <BookIcon className="Icon" />,
    movie: <MovieIcon className="Icon" />,
    video_game: <VideoGameIcon className="Icon" />,
    album: <AlbumIcon className="Icon" />,
  };

  const dispatch = useDispatch();

  const payload = {
    type,
    data: { title, author, pub_year, page_count, blockID },
  };

  //on mount, populate the database data (in the state) with the block information
  useEffect(() => {
    dispatch(populateDatabaseData(payload));

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
        addImageToDatabaseData({
          type,
          src,
          idx,
          blockID,
        })
      );
    } else {
      dispatch(removeImageFromDatabaseData({ blockID, type, idx }));
    }
  };

  //get genres for checkbox population
  const genres = useContext(GenreContext);

  return (
    <div className={`Block ${type}`}>
      {icons[type]}
      <div className="imageContainer">
        {images.map((src, idx) => (
          <div
            className="imageWrapper"
            key={src}
            onClick={() => handleClick(blockID, type, idx, src)}
          >
            <img className={`${type}-img `} src={src}></img>
            <div className={`overlay ${clicked[idx] ? "show" : ""}`}>
              <p>Selected</p>
            </div>
          </div>
        ))}
      </div>
      <div className="titleInfoContainer">
        <MyTextArea name="title" label="Title" />
        {type === "book" ? (
          <>
            <MyTextArea name="author" label="Author" />
            <MyTextArea name="pub_year" label="Publication Year" />
            <MyTextArea name="page_count" label="Page Count" />
          </>
        ) : (
          <></>
        )}
      </div>
      {type === "book" ? (
        <div className="genreCheckboxes">
          {genres?.map((text, idx) => (
            <label key={idx} className="MCC-font">
              <input type="checkbox" />
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
