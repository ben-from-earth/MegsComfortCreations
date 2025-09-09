import "./LoadingWidget.css";
import CircularProgress from "@mui/material/CircularProgress";

const LoadingWidget = ({ message }) => {
  return (
    <div className="LoadingWidget MCC-font">
      <p>{message}</p>
      <CircularProgress sx={{ color: "#e1b3b5" }} />
    </div>
  );
};

export default LoadingWidget;
