import "./CollectedCoversBlock.css";
import { v4 as uuid } from "uuid";
import BookIcon from "@mui/icons-material/BookTwoTone";
import MovieIcon from "@mui/icons-material/LocalMoviesTwoTone";
import VideoGameIcon from "@mui/icons-material/VideogameAssetTwoTone";
import AlbumIcon from "@mui/icons-material/AlbumTwoTone";

const CollectedCoversBlock = ({
  info: {
    type,
    images,
    blockInfo: { title, author, first_publish_year, number_of_pages },
  },
}) => {
  const icons = {
    Book: <BookIcon className="Icon" />,
    Movie: <MovieIcon className="Icon" />,
    "Video Game": <VideoGameIcon className="Icon" />,
    Album: <AlbumIcon className="Icon" />,
  };

  return (
    <div className="Block">
      {icons[type]}
      <div className="imageContainer">
        {images.map((i) => (
          <img key={uuid()} src={i}></img>
        ))}
      </div>
      <div className="titleInfoContainer">
        <label className="MCC-font" htmlFor="title">
          Title:
        </label>
        <textarea name="title" defaultValue={title}></textarea>
        {author ? (
          <>
            <label className="MCC-font" htmlFor="author">
              Author:
            </label>
            <textarea name="author" defaultValue={author}></textarea>
            <label className="MCC-font" htmlFor="pubYear">
              Publication Year:
            </label>
            <textarea
              name="pubYear"
              defaultValue={first_publish_year ? first_publish_year : ""}
            ></textarea>
            <label className="MCC-font" htmlFor="pageCount">
              Page Count:
            </label>
            <textarea
              name="pageCount"
              defaultValue={number_of_pages ? number_of_pages : ""}
            ></textarea>
          </>
        ) : (
          <></>
        )}
      </div>
    </div>
  );
};

export default CollectedCoversBlock;
