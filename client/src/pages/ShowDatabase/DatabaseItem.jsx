const DatabaseItem = ({
  info: { title, author, page_count, pub_year, spine_color, image_urls },
  type,
}) => {
  return (
    <div className={`databaseItem ${type}`}>
      {image_urls.map((src, idx) => (
        <img
          id={idx}
          className={`databaseDisplayImage-${type}`}
          src={src}
        ></img>
      ))}
      {type === "book" ? (
        <p>
          {title} by {author} - {page_count} pages - {pub_year}
        </p>
      ) : (
        <p>{title}</p>
      )}
    </div>
  );
};

export default DatabaseItem;
