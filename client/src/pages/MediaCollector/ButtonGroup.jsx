import "./ButtonGroup.css";
import { useDispatch, useSelector } from "react-redux";
import {
  selectDatabaseData,
  sendToDatabase,
} from "../../state/databaseDataSlice";

const handleDatabaseClick = async (dispatch, databaseData) => {
  const responses = await dispatch(sendToDatabase({ databaseData }));
  console.log(responses.payload); //capturing database creation responses here. Keeping as log until handling it.
};

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
        onClick={() => handleDatabaseClick(dispatch, databaseData)}
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
