import { titleRearrange } from "../MediaCollector/helpers/mediaCollectorHelpers";

const DatabaseItem = ({
  info: { title, author, page_count, pub_year, spine_color, image_urls },
  type,
}) => {
  //classes based on type
  const typeClasses = {
    book: "bg-[#98ab88] border-[#3d770d]",
    movie: "bg-[#323b43] border-black text-white",
    album: "bg-[#7fa5a3] border-[#d49a97]",
    video_game: "bg-[#98ab88] border-[#4e8885]",
  };

  return (
    <div
      className={`mr-auto box-border flex w-full items-center justify-start gap-5 rounded-sm border-2 p-2 ${typeClasses[type]}`}
    >
      {image_urls.map((src, idx) => (
        <img
          id={idx}
          className={type === "album" ? "h-[75px]" : "w-15"}
          src={src}
        ></img>
      ))}
      {type === "book" ? (
        <p>
          {titleRearrange(title)} by {author} - {page_count} pages - {pub_year}
        </p>
      ) : (
        <p>{titleRearrange(title)}</p>
      )}
    </div>
  );
};

export default DatabaseItem;
