import { useContext } from "react";
import "./ButtonGroup.css";
import DataContext from "./DataContext";
import MediaDataContext from "./MediaDataContext";

const ButtonGroup = () => {
  const { dispatch } = useContext(DataContext);
  const { CollectedCoversBlocks } = useContext(MediaDataContext);
  return (
    <div className="ButtonGroup">
      <button
        className="MCC-font"
        onClick={() => dispatch({ type: "Collect" })}
      >
        Collect Media Covers
      </button>
      <button
        className="MCC-font"
        onClick={() =>
          dispatch({ type: "send-to-database", items: CollectedCoversBlocks })
        }
      >
        Send to Database
      </button>
    </div>
  );
};

export default ButtonGroup;
