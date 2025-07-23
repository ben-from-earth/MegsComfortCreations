import { useContext } from "react";
import "./ButtonGroup.css";
import DataContext from "./DataContext";

const ButtonGroup = () => {
  const { dispatch } = useContext(DataContext);
  return (
    <div className="ButtonGroup">
      <button onClick={() => dispatch({ type: "Collect" })}>
        Collect Media Covers
      </button>
      <button>Send to Database</button>
    </div>
  );
};

export default ButtonGroup;
