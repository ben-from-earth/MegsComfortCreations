import "./DatabaseSavedWidget.css";

const DatabaseSavedWidget = ({ data, close }) => {
  return (
    <div className="DatabaseSavedWidget MCC-font">
      Successfully added the following items to the database:
      <ol>
        {data.map((item, idx) => {
          console.log(item);
          let listItem;
          let keys = Object.keys(item);
          console.log(keys);
          let type = keys
            .filter((key) => key.startsWith("saved_"))[0]
            .split("_")[1];
          if (type === "book") {
            listItem = `(${type}) ${item.saved_book.title} by ${item.saved_book.author}`;
          } else {
            listItem = `(${type}) ${item[`saved_${type}`].title}`;
          }
          return <li key={idx}>{listItem}</li>;
        })}
      </ol>
      <button onClick={close} className="MCC-font">
        Close
      </button>
    </div>
  );
};

export default DatabaseSavedWidget;
