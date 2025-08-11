import "./ButtonGroup.css";
import { useDispatch } from "react-redux";
import { collectMedia, setCollectText } from "../../app/collectorSlice";

const ButtonGroup = ({ mediaTypes, searchData, setSearchData }) => {
  const dispatch = useDispatch();
  return (
    <div className="ButtonGroup">
      <button
        className="MCC-font"
        onClick={() => {
          dispatch(setCollectText({ searchData }));
          dispatch(collectMedia());
          setSearchData(
            mediaTypes.map((media) => ({ type: media.type, text: "" }))
          );
        }}
      >
        Collect Media Covers
      </button>
      <button
        className="MCC-font"
        // onClick={() =>
        //   dispatch({ type: "send-to-database", items: CollectedCoversBlocks })
        // }
      >
        Send to Database
      </button>
    </div>
  );
};

export default ButtonGroup;
