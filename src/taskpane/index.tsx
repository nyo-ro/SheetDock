import { createRoot } from "react-dom/client";
import { App } from "./App";
const root = document.getElementById("root");
if (!root) throw new Error("SheetDockの表示領域がありません。");
createRoot(root).render(<App />);
if (process.env.NODE_ENV === "development") {
  const experimentalRoot = document.createElement("div");
  root.after(experimentalRoot);
  void import("../poc/Panel").then(({ Panel }) => createRoot(experimentalRoot).render(<Panel />)).catch(error => {
    experimentalRoot.textContent = "検証UIの読み込みに失敗しました: " + String(error);
  });
}
