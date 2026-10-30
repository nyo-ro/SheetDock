import { createRoot } from "react-dom/client";
import { App } from "./App";
const root = document.getElementById("root");
if (!root) throw new Error("SheetDockの表示領域がありません。");
const renderNormal = () => createRoot(root).render(<App />);
if (process.env.NODE_ENV === "development") {
  // 実験モードの判定を先に行い、通常UI・PoCのeffectを起動しない。
  void import("../dev/bootstrap").then(({ startDevelopment }) => startDevelopment(root, renderNormal)).catch(() => {
    root.textContent = "開発画面の読み込みに失敗しました。ページを再読み込みしてください。";
  });
} else {
  renderNormal();
}
