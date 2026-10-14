import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./components/App";

Office.onReady(() => {
  const container = document.getElementById("container");
  if (!container) {
    return;
  }
  createRoot(container).render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>,
  );
});
