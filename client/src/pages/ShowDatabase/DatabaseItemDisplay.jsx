import DatabaseItem from "./DatabaseItem";

const DatabaseItemDisplay = ({ databaseItems, page, limit, total }) => {
  return (
    <div className="databaseItemDisplay MCC-font">
      <p>This is the item display</p>
      {databaseItems.items.map((item) => {
        return (
          <DatabaseItem key={item.id} info={item} type={databaseItems.type} />
        );
      })}
    </div>
  );
};

export default DatabaseItemDisplay;
