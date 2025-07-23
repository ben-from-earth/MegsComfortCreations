import BackgroundIMG from "./assets/FlowerBackground.png";
import MainMenuTitle from "./assets/MegsMediaCollector.png";
import "./MainMenu.css";

const MainMenu = () => {
  return (
    <div
      className="InfoForm"
      style={{
        backgroundImage: `url(${BackgroundIMG})`,
      }}
    >
      <img src={`${MainMenuTitle}`} />
      <button>Gather Media Images</button>
      <button>Go to Database</button>
      <button>Edit Metadata</button>
      <button>Search by Metadata</button>
    </div>
  );
};

export default MainMenu;
