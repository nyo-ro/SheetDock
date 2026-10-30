import { useEffect, useState } from "react";
import { ensureExcelReady } from "../../excel/environment";
import { prepareKnownSheets, ProbeKind, runReadProbe } from "./api";
import { ReadProbeController } from "./controller";
import { reloadInMode } from "./mode";
import "./probe.css";

const phaseText = { idle: "未実行", preparing: "Bの既知IDを準備中", ready: "Bの準備完了", scheduled: "予約待ち", running: "API応答待ち", succeeded: "取得成功", failed: "取得失敗", stalled: "15秒以上応答待ち", settledLate: "期限超過後に応答（試験終了）", cancelled: "予約取消（試験終了）", disposed: "実行環境終了" };
export function ReadProbePanel() {
  const [controller] = useState(() => new ReadProbeController({ ready: ensureExcelReady, prepare: prepareKnownSheets, run: runReadProbe }));
  const [, redraw] = useState(0);
  const [kind, setKind] = useState<ProbeKind>("A");
  const [target, setTarget] = useState("");
  const [delay, setDelay] = useState("5");
  const [exported, setExported] = useState("");
  const [navigationError, setNavigationError] = useState(false);
  useEffect(() => {
    const unsubscribe = controller.subscribe(() => redraw(n => n + 1));
    // 表示時計だけ。Excel APIは呼ばない。
    const timer = setInterval(() => redraw(n => n + 1), 250);
    const dispose = () => controller.dispose();
    window.addEventListener("pagehide", dispose);
    return () => { unsubscribe(); clearInterval(timer); window.removeEventListener("pagehide", dispose); controller.dispose(); };
  }, [controller]);
  const view = controller.view();
  const diagnostic = controller.diagnostics();
  const delayMs = Number(delay) * 1000;
  const delayValid = delay.trim() !== "" && Number.isInteger(delayMs) && delayMs >= 0 && delayMs <= 30000;
  const chosen = target || view.choices[0]?.id;
  const canSchedule = delayValid && (kind === "B" ? view.phase === "ready" && !!chosen : view.phase === "idle");
  const canReload = !view.pending && view.phase !== "scheduled";
  function navigate(experiment: boolean) {
    if (!canReload) return;
    try { reloadInMode(experiment); } catch { setNavigationError(true); }
  }
  return <main className="read-probe">
    <h1>読み取り比較 A・B・C</h1>
    <p>開発専用・読み取りのみ。通常UI・PoC・イベント購読は起動していません。</p>
    <p>1ページ起動につき1試験です。別の試験は新しい実行環境で行います。</p>
    <p>実行環境: <code>{controller.session}</code></p>
    <fieldset disabled={view.phase !== "idle"}>
      <legend>比較する取得処理</legend>
      <label><input type="radio" name="probe-kind" checked={kind === "A"} onChange={() => setKind("A")} />A：ワークシート一覧のみ</label>
      <label><input type="radio" name="probe-kind" checked={kind === "B"} onChange={() => setKind("B")} />B：既知IDのワークシート</label>
      <label><input type="radio" name="probe-kind" checked={kind === "C"} onChange={() => setKind("C")} />C：アクティブシートIDのみ</label>
    </fieldset>
    {kind === "B" && <section>
      <p>Bのみ、名前編集を始める前に1回の一覧取得が必要です。準備と本試験は別々に記録します。</p>
      <button disabled={view.phase !== "idle"} onClick={() => void controller.prepare()}>Bの対象を準備</button>
      <label>対象シート<select aria-label="対象シート" value={chosen ?? ""} disabled={view.phase !== "ready"} onChange={e => setTarget(e.target.value)}>
        {!view.choices.length && <option value="">準備後に選択</option>}
        {view.choices.map(s => <option key={s.id} value={s.id}>{s.position + 1}：{s.name}</option>)}
      </select></label>
    </section>}
    <label>開始までの秒数（0〜30）<input aria-label="開始までの秒数" type="number" min="0" max="30" step="1" value={delay} disabled={view.phase !== "idle" && view.phase !== "ready"} onChange={e => setDelay(e.target.value)} /></label>
    <button disabled={!canSchedule} onClick={() => controller.schedule(kind, kind === "B" ? chosen : undefined, delayMs)}>読み取りを予約実行</button>
    {view.phase === "scheduled" && <><p role="status">API開始まで約 {Math.max(0, ((view.dueAt! - Date.now()) / 1000)).toFixed(1)} 秒</p><button onClick={() => controller.cancelSchedule()}>予約を取り消す</button></>}
    <p role="status">状態: {phaseText[view.phase]}{view.count !== undefined ? ` ／取得件数: ${view.count}` : ""}</p>
    {view.waitingAt && <p>待機箇所: <code>{view.waitingAt}</code></p>}
    {view.phase === "failed" && <p role="alert">取得に失敗しました。コード: {view.errorCode}。診断ログを保存し、新しい実行環境で次の試験を行ってください。</p>}
    {view.phase === "stalled" && <p role="alert">15秒以上待機しています。処理は取り消していません。この環境で次のAPIは開始できません。ログを保存してExcelのブラウザータブ全体を再読み込みしてください。</p>}
    {view.phase === "settledLate" && <p role="alert">遅れて処理が終了しました。再試行は行いません。結果を保存して新しい実行環境へ進んでください。</p>}
    <p>本試験: Excel.run {diagnostic.apiCalls.experiment.run} 回 / sync {diagnostic.apiCalls.experiment.sync} 回<br />準備: Excel.run {diagnostic.apiCalls.preparation.run} 回 / sync {diagnostic.apiCalls.preparation.sync} 回</p>
    <section>
      <h2>操作時刻の手動記録</h2>
      <p>記録するのはボタンを押した時刻です。Excel側の編集開始やEscを自動検知しません。ボタンへの移動でフォーカスが変わるため、実際の操作時刻とは差があります。</p>
      <button onClick={() => controller.mark("editStart")}>名前編集開始を記録（手動）</button>
      <button onClick={() => controller.mark("escape")}>Esc取消を記録（手動）</button>
    </section>
    <section>
      <h2>診断ログ</h2>
      <p>シート名・ID・ブック内容・URL・生のエラー本文は含めません。時刻はUTCと実行環境開始からの経過時間です。</p>
      <button onClick={() => setExported(JSON.stringify(controller.diagnostics(), null, 2))}>診断ログを表示・更新</button>
      {exported && <label>このJSONを全選択してコピー<textarea aria-label="診断ログJSON" readOnly value={exported} rows={12} /></label>}
      <ol className="probe-log">{diagnostic.entries.slice(-10).map(row => <li key={row.seq}>{row.elapsedMs}ms {row.stage}{row.durationMs !== undefined ? ` (${row.durationMs}ms)` : ""}</li>)}</ol>
    </section>
    <button disabled={!canReload} onClick={() => navigate(true)}>新しい実行環境へ（再読み込み）</button>
    <button disabled={!canReload} onClick={() => navigate(false)}>通常画面へ（再読み込み）</button>
    {navigationError && <p role="alert">モード保存・再読み込みに失敗しました。ブラウザーのストレージ設定を確認してください。</p>}
    <p>未応答中に画面内で別試験へ切り替えることはできません。Excelのブラウザータブ全体を再読み込みした後、実行環境の値が変わったことを確認してください。</p>
  </main>;
}
