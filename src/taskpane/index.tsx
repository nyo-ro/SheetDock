import { createRoot } from "react-dom/client";
import { App } from "./App";
const root = document.getElementById("root");
if (!root) throw new Error("SheetDockの表示領域がありません。");
createRoot(root).render(<App />);
