import BackgroundIMG from "./assets/FlowerBackground.png";
import MainMenuTitle from "./assets/MegsMediaCollector.png";
import { useReducer } from "react";
import "./MainMenu.css";
import MediaInputs from "./MediaInputs";
import ButtonGroup from "./ButtonGroup";
import DataContext from "./DataContext";

function reducer(state, action) {
  switch (action.type) {
    case "set-checks":
      return state.map((mediaType, i) =>
        action.idx === i ? { ...mediaType, show: !mediaType.show } : mediaType
      );
    case "set-collect-text": {
      let searchArr = action.text.split("/").map((i) => i.trim());
      searchArr = searchArr.map((t) =>
        t
          ? t +
            ` ${action.mediaType.slice(
              0,
              action.mediaType.length - 1
            )} Cover Image`
          : ""
      );
      return state.map((mediaType) =>
        action.mediaType === mediaType.mediaType
          ? { ...mediaType, titles: searchArr }
          : mediaType
      );
    }
    case "Collect": {
      console.log(state);
      return state;
    }
  }
}

const MainMenu = () => {
  const medias = ["Books", "Movies", "Video Games", "Albums"];
  const [Data, dispatch] = useReducer(
    reducer,
    medias.map((m) => ({ mediaType: m, show: false, titles: [] }))
  );

  return (
    <DataContext.Provider value={{ dispatch }}>
      <div
        className="InfoForm"
        style={{
          backgroundImage: `url(${BackgroundIMG})`,
        }}
      >
        <img src={`${MainMenuTitle}`} />
        <MediaInputs info={Data} />
        <ButtonGroup />
      </div>
    </DataContext.Provider>
  );
};

export default MainMenu;
