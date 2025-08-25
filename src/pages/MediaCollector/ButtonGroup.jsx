import "./ButtonGroup.css";
import { useDispatch, useSelector } from "react-redux";
import {
  selectDatabaseData,
  sendToDatabase,
} from "../../state/databaseDataSlice";

const ButtonGroup = ({ onCollect, onPNG }) => {
  // setup connection to redux slice and get all searched information
  const dispatch = useDispatch();
  const databaseData = useSelector(selectDatabaseData);

  return (
    <div className="ButtonGroup">
      <button
        className="MCC-font"
        onClick={() => {
          onCollect();
        }}
      >
        Collect Media Covers
      </button>
      <button
        className="MCC-font"
        onClick={() => dispatch(sendToDatabase({ databaseData }))}
      >
        Send to Database
      </button>
      <button className="MCC-font" onClick={() => onPNG()}>
        Get PNG
      </button>
    </div>
  );
};

export default ButtonGroup;
