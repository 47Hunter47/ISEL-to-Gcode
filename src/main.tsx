import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import "./styles.css";

// the app is dark-only; shadcn's tokens switch on this class
window.document.documentElement.classList.add("dark");

// desktop feel: no browser context menu (Ctrl+C/V and the log's select-text
// still cover copy/paste)
document.addEventListener("contextmenu", (e) => e.preventDefault());

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
