import { createRoot } from "react-dom/client";
import { useState } from "react";
import { isReadProbeMode, reloadInMode } from "./readProbe/mode";
function ProbeLauncher() {
  const [failed, setFailed] = useState(false);
  return <section className="probe-launcher"><button onClick={() => { try { reloadInMode(true); } catch { setFailed(true); } }}>読み取り比較モードを開く（再読み込み）</button>
    <p>通常UIとPoCを起動しない、A・B・Cの読み取り専用実験です。</p>
    {failed && <p role="alert">モードを保存できません。ブラウザーのストレージ設定を確認してください。</p>}
  </section>;
}
export async function startDevelopment(root: HTMLElement, renderNormal: () => void): Promise<void> {
  if (isReadProbeMode()) {
    const { ReadProbePanel } = await import("./readProbe/Panel");
    createRoot(root).render(<ReadProbePanel />);
    return;
  }
  renderNormal();
  const launcher = document.createElement("div"); root.after(launcher); createRoot(launcher).render(<ProbeLauncher />);
  const experimentalRoot = document.createElement("div"); launcher.after(experimentalRoot);
  try { const { Panel } = await import("../poc/Panel"); createRoot(experimentalRoot).render(<Panel />); }
  catch { experimentalRoot.textContent = "検証UIの読み込みに失敗しました。"; }
}
