import { useState } from "react";
import DatabaseItemDisplay from "./DatabaseItemDisplay";
import PaginationInputs from "./PaginationInputs";

const ShowDatabasePage = () => {
  const [databaseItems, setDatabaseItems] = useState({ type: "", items: [] });
  return (
    <div className="flex flex-col items-center">
      <PaginationInputs setDatabaseItems={setDatabaseItems} />
      <DatabaseItemDisplay databaseItems={databaseItems} />
    </div>
  );
};

export default ShowDatabasePage;
