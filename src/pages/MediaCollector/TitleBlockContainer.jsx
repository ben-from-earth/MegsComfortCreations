import CollectedCoversBlock from "./CollectedCoversBlock";
import "./TitleBlockContainer.css";
import { memo } from "react";

const TitleBlockContainer = memo(function ({ blocks }) {
  return (
    <div className="TitleBlockContainer">
      {blocks.map((b) => (
        <CollectedCoversBlock info={b} key={b.blockID} />
      ))}
    </div>
  );
});

export default TitleBlockContainer;
