import "./TitleBlockContainer.css";

const TitleBlockContainer = ({ blocks }) => {
  return <div className="TitleBlockContainer">{blocks.map((t) => t)}</div>;
};

export default TitleBlockContainer;
