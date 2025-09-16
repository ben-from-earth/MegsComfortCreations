import { useSelector } from "react-redux";
import { selectDatabaseData } from "@/state/databaseDataSlice";
import Button from "@/components/Button";

const ButtonGroup = ({ onCollect, onPNG, onDatabase }) => {
  // setup connection to redux slice and get the database information
  const databaseData = useSelector(selectDatabaseData);

  return (
    <div className="flex flex-row content-center gap-4">
      <Button
        onClick={() => {
          onCollect();
        }}
        label={"Collect Media Covers"}
        width={175}
        fontSize={25}
      />
      <Button
        onClick={() => onDatabase(databaseData)}
        label={"Send to Database"}
        width={175}
        fontSize={25}
      />
      <Button
        onClick={() => onPNG()}
        label={"Get PNG"}
        width={175}
        fontSize={25}
      />
    </div>
  );
};

export default ButtonGroup;
