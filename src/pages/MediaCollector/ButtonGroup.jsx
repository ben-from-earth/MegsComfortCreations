import "./ButtonGroup.css";
import { useDispatch, useSelector } from "react-redux";
import {
  selectDatabaseData,
  sendToDatabase,
} from "../../app/databaseDataSlice";

const ButtonGroup = ({ onCollect }) => {
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
    </div>
  );
};

export default ButtonGroup;
