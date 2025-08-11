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
  updateDatabaseData,
} from "../../app/databaseDataSlice";
import { memo, useEffect, useState } from "react";

const CollectedCoversBlock = memo(function CollectedCoversBlock({
  info: {
    type,
    images,
    blockInfo: { title, author, first_publish_year, number_of_pages },
    id,
  },
}) {
  const MyTextArea = ({ name, label }) => {
    const dispatch = useDispatch();

    const value = useSelector((s) => {
      // databaseData: [{ id, label, data: [...] }]
      // data: [{ title, blockID, images, ... }]
      const group = s.databaseData.find((g) => g.id === type);
      const block = group?.data?.find((d) => d.blockID === id);
      // Use store value if present; otherwise fallback from props
      return block ? block[name] : "";
    });

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
              updateDatabaseData({ id, type, name, newText: e.target.value })
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
    data: { title, author, first_publish_year, number_of_pages, blockID: id },
  };
  useEffect(() => {
    dispatch(populateDatabaseData(payload));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dispatch]);

  const [clicked, setClicked] = useState(() =>
    Array(images.length).fill(false)
  );

  const handleClick = (id, type, idx) => {
    const next = !clicked[idx];
    setClicked((prev) =>
      prev.map((b, itemIndex) => (itemIndex === idx ? next : b))
    );
    if (next) {
      dispatch(
        addImageToDatabaseData({
          type,
          text: `image${idx + 1}`,
          idx,
          id,
        })
      );
    } else {
      dispatch(removeImageFromDatabaseData({ id, type, idx }));
    }
  };

  return (
    <div className="Block">
      {icons[type]}
      <div className="imageContainer">
        {images.map((src, idx) => (
          <div
            className="imageWrapper"
            key={src}
            onClick={() => handleClick(id, type, idx)}
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
            <MyTextArea name="first_publish_year" label="Publication Year" />
            <MyTextArea name="number_of_pages" label="Page Count" />
          </>
        ) : (
          <></>
        )}
      </div>
    </div>
  );
});

export default CollectedCoversBlock;
