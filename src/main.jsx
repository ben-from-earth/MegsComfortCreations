import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { store } from "./app/store";
import { Provider } from "react-redux";
import "./index.css";
import App from "./App.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <Provider store={store}>
      <App />
    </Provider>
  </StrictMode>
);
