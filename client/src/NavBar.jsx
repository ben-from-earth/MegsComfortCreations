import logo from "./assets/Comfort.png";
import { NavLink } from "react-router";
import "./NavBar.css";

const NavBar = () => {
  return (
    <nav className="NavBar">
      <img src={logo} />

      <div className="linkGroup">
        <NavLink to="/" className="MCC-font">
          Home
        </NavLink>

        <NavLink to="Shop" className="MCC-font">
          Shop
        </NavLink>

        <NavLink to="MegsRecs" className="MCC-font">
          Meg's Recs
        </NavLink>

        <NavLink to="Newsletter" className="MCC-font">
          Newsletter
        </NavLink>
        <NavLink to="ShowDatabase" className="MCC-font">
          Show Database
        </NavLink>

        <NavLink to="MediaCollector" className="MCC-font">
          Media Collector
        </NavLink>
      </div>
    </nav>
  );
};

export default NavBar;
