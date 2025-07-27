import "./CollectedCoversBlock.css";
import { v4 as uuid } from "uuid";
import BookIcon from "@mui/icons-material/BookTwoTone";
import MovieIcon from "@mui/icons-material/LocalMoviesTwoTone";
import VideoGameIcon from "@mui/icons-material/VideogameAssetTwoTone";
import AlbumIcon from "@mui/icons-material/AlbumTwoTone";

const CollectedCoversBlock = ({
  type,
  images,
  blockInfo: { title, author, first_publish_year, number_of_pages },
}) => {
  const icons = {
    Book: <BookIcon />,
    Movie: <MovieIcon />,
    VideoGame: <VideoGameIcon />,
    Album: <AlbumIcon />,
  };

  return (
    <div className="Block">
      {icons[type]}
      <div className="imageContainer">
        {images.map((i) => (
          <img key={uuid()} src={i}></img>
        ))}
      </div>
      <p>Title: {title}</p>
      {author ? (
        <>
          <p>Author: {author}</p>
          <p>Publication Year: {first_publish_year}</p>
          <p>Page Count: {number_of_pages}</p>
        </>
      ) : (
        <></>
      )}
    </div>
  );
};

export default CollectedCoversBlock;
