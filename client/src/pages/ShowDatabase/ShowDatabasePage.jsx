import { useState } from "react";
import DatabaseItemDisplay from "./DatabaseItemDisplay";
import PaginationInputs from "./PaginationInputs";
import "./ShowDatabasePage.css";

const ShowDatabasePage = () => {
  const [databaseItems, setDatabaseItems] = useState({ type: "", items: [] });
  return (
    <div className="ShowDatabasePage">
      <PaginationInputs setDatabaseItems={setDatabaseItems} />
      <DatabaseItemDisplay databaseItems={databaseItems} />
    </div>
  );
};

export default ShowDatabasePage;
