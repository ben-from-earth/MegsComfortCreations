import "./LoadingWidget.css";
import CircularProgress from "@mui/material/CircularProgress";

const LoadingWidget = ({ searchCount }) => {
  return (
    <div className="LoadingWidget MCC-font">
      <p>Getting {searchCount} Media Covers</p>
      <CircularProgress sx={{ color: "#e1b3b5" }} />
    </div>
  );
};

export default LoadingWidget;
