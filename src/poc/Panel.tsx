import { diagnose } from "../diagnostics";
import { useEffect, useRef, useState } from "react";
import { ensureExcelReady } from "../excel/environment";
import { withExcelLock } from "../excel/serial";
import { publishSnapshot } from "../state/snapshotBus";
import { CoalescedRefresh } from "../sync/coalescedRefresh";
import { groupOrder, positionPlan } from "../domain/movePlan";
import { Action, BookSnapshot, describeError } from "./model";
import { applyAction, observeBook } from "./officeAdapter";
import { runExperiment } from "./experiment";
import { EventRecord, eventLifecycle, eventSets, listenEvents } from "./events";
import "./poc.css";
import { ConfirmationDialog, OperationPlan } from "./ConfirmationDialog";
export function Panel() {
  const [enabled,setEnabled] = useState(false);
  const [onlyWorksheets,setOnlyWorksheets] = useState(false);
  const [bookName,setBookName] = useState("");
  const [snapshot,setSnapshot] = useState<BookSnapshot>();
  const [events,setEvents] = useState<EventRecord[]>([]);
  const [auto,setAuto] = useState(false);
  const [busy,setBusy] = useState(false);
  const [target,setTarget] = useState("");
  const [kind,setKind] = useState("Hidden");
  const [boundary,setBoundary] = useState("0");
  const [nameProbe,setNameProbe] = useState("");
  const [failureStep,setFailureStep] = useState("0");
  const [failureWhen,setFailureWhen] = useState<"before"|"after">("before");
  const [plan,setPlan] = useState<OperationPlan>();
  const [confirmation,setConfirmation] = useState<OperationPlan>();
  const pendingConfirmation = useRef<OperationPlan>();
  const [output,setOutput] = useState<unknown>();
  const [fault,setFault] = useState<string>();
  const [refreshCount,setRefreshCount] = useState(0);
  const alive = useRef(true); const lock = useRef(false);
  const scheduler = useRef<CoalescedRefresh>();
  async function refresh() {
    await ensureExcelReady();
    if(!alive.current)return;
    await withExcelLock(async () => {
      if(!alive.current)return;
      const current = await observeBook();
      if (!alive.current) return;
      setSnapshot(current); setRefreshCount(count => count+1);
      // 観測だけでは通常UIに再取得を要求しない。各UIが同じイベントを購読する。
    });
  }
  useEffect(() => {
    alive.current = true;diagnose("poc.init");
    const queue = new CoalescedRefresh(refresh, error => {
      if (alive.current) { setFault(JSON.stringify(describeError(error))); }
    },150,"poc.refresh");
    scheduler.current = queue;
    const removeListener = listenEvents(record => { setEvents(rows => [...rows.slice(-149),record]); queue.request(); });
    const focus = () => { if (document.visibilityState === "visible") queue.request(); };
    // フォーカス復帰は専用イベントのない変更への読み取りフォールバック。
    const enabledFocus = () => { if (autoRef.current) focus(); };
    window.addEventListener("focus",enabledFocus); document.addEventListener("visibilitychange",enabledFocus);
    return () => {
      alive.current = false;diagnose("poc.dispose"); queue.stop(); removeListener(); window.removeEventListener("focus",enabledFocus); document.removeEventListener("visibilitychange",enabledFocus);
      void eventLifecycle.setEnabled(false).catch(error => console.error("SheetDock PoC event cleanup failed",error));
    };
  }, []);
  const autoRef = useRef(false);
  async function task(work: () => Promise<void>) {
    if (lock.current) return; lock.current = true; setBusy(true); setFault(undefined);
    try { await work(); } catch (error) { if (alive.current) setFault(JSON.stringify(describeError(error))); }
    finally { lock.current = false; if (alive.current) setBusy(false); }
  }
  async function toggleAuto(value: boolean) {
    await ensureExcelReady(); if(!alive.current)return;
    await eventLifecycle.setEnabled(value);
    if(!alive.current){await eventLifecycle.setEnabled(false);return;}
    autoRef.current = value; setAuto(value); if (value) scheduler.current?.request();
  }
  function prepare() {
    try {
      if (!snapshot) throw new Error("先に観測してください。");
      const ids = target.split(/[,\n]/).map(s => s.trim()).filter(Boolean);
      if (!ids.length || new Set(ids).size !== ids.length || ids.some(id => !snapshot.sheets.some(s => s.id === id))) throw new Error("一覧にある重複しないIDを指定してください。");
      let actions: Action[] = []; let expectedOrder: string[] | undefined;
      if (kind === "move") {
        const full = snapshot.sheets.map(s => s.id);
        expectedOrder = groupOrder(full,full,ids,Number(boundary));
        actions = positionPlan(full,ids,expectedOrder).map(write => ({kind:"move",...write}));
      } else {
        if (ids.length !== 1) throw new Error("移動以外のPoCは対象を1つ指定してください。");
        if (kind === "deleteVeryHidden") {
          if (snapshot.sheets.find(s => s.id === ids[0])?.visibility !== "VeryHidden") throw new Error("VeryHiddenの対象を指定してください。");
          actions = [{kind:"visibility",id:ids[0],visibility:"Hidden"},{kind:"delete",id:ids[0]}];
        } else if (kind === "copy" || kind === "delete") actions = [{kind,id:ids[0]}];
        else if (kind === "Visible" || kind === "Hidden" || kind === "VeryHidden") actions = [{kind:"visibility",id:ids[0],visibility:kind}];
      }
      const step = Number(failureStep);
      if (!Number.isInteger(step) || step < 0 || step > actions.length) throw new Error("失敗注入番号は0または計画内のステップ番号にしてください。");
      setPlan({baseline:structuredClone(snapshot),actions,options:{worksheetsOnly:onlyWorksheets,expectedOrder,
        veryHiddenDeleteId:kind === "deleteVeryHidden" ? ids[0] : undefined,
        injectFailure:step ? {step,when:failureWhen} : undefined}});
      setFault(undefined);
    } catch (error) { setFault(JSON.stringify(describeError(error))); }
  }
  function cancelConfirmation() {
    pendingConfirmation.current = undefined;
    setConfirmation(undefined);
  }
  function requestConfirmation() {
    if (!canRun || lock.current || pendingConfirmation.current || !plan) return;
    const frozenPlan = structuredClone(plan);
    pendingConfirmation.current = frozenPlan;
    setConfirmation(frozenPlan);
  }
  function confirmExecution() {
    const selectedPlan = pendingConfirmation.current;
    if (!selectedPlan || lock.current) return;
    // Reactの再描画より前に確認を消費して、連打や古いイベントからの再実行を防ぐ。
    cancelConfirmation();
    void task(() => execute(selectedPlan));
  }
  async function execute(selectedPlan: OperationPlan) {
    if (!enabled || !onlyWorksheets || bookName !== selectedPlan.baseline.name) throw new Error("検証ブックの有効化・種類確認・ブック名の一致が必要です。");
    setPlan(undefined); setOutput(undefined);
    await ensureExcelReady();
    await withExcelLock(async () => {
      const result = await runExperiment({read:observeBook,apply:applyAction},selectedPlan.baseline,selectedPlan.actions,{...selectedPlan.options,worksheetsOnly:onlyWorksheets});
      if (alive.current) {
        setOutput(result); setSnapshot(result.after);
        publishSnapshot({sheets:result.after?.activeId ? result.after.sheets : [],error:result.after?.activeId ? undefined : "操作後の状態が不明です。手動更新してください。"});
      }
    });
  }
  const canRun = !!plan && enabled && onlyWorksheets && bookName === plan.baseline.name && !busy && !confirmation;
  return <section className="poc-panel" aria-label="Office.js技術検証">
    <h2>Office.js技術検証（開発専用）</h2>
    <p>専用の使い捨て検証ブックのみで使用してください。実機検証完了・安全性保証を意味しません。</p>
    <button disabled={busy || !!confirmation} onClick={() => void task(refresh)}>観測・再取得（変更なし）</button>
    <label><input type="checkbox" checked={auto} disabled={busy || !!confirmation} onChange={e => void task(() => toggleAuto(e.target.checked))} />イベント自動同期を検証</label>
    <p>取得回数: {refreshCount} ／履歴は直近150件。ブック構成保護は操作前・手動・フォーカス復帰で再取得。</p>
    <details><summary>対応API情報</summary><pre>{JSON.stringify({eventSets,required:"ExcelApi 1.17",supported:typeof Office !== "undefined" && !!Office.context?.requirements?.isSetSupported("ExcelApi","1.17")},null,2)}</pre></details>
    <details open><summary>ワークシート観測値</summary><pre>{JSON.stringify(snapshot,null,2) ?? "未取得"}</pre></details>
    <label>名前重複の観測<input value={nameProbe} onChange={e => setNameProbe(e.target.value)} /></label>
    <p>{nameProbe && snapshot ? (snapshot.sheets.some(s => s.name.toLowerCase() === nameProbe.toLowerCase()) ? "通常ワークシート内に同名あり" : "通常ワークシート内に一致なし。グラフシートとの重複・名前の利用可否は不明。") : "名前チェックは読み取り専用です。"}</p>
    <fieldset disabled={busy || !!confirmation}>
      <legend>変更操作の明示的な有効化</legend>
      <label><input type="checkbox" checked={enabled} onChange={e => {setEnabled(e.target.checked);setPlan(undefined);}} />使い捨て検証ブックで変更操作を行う</label>
      <label><input type="checkbox" checked={onlyWorksheets} onChange={e => {setOnlyWorksheets(e.target.checked);setPlan(undefined);}} />通常ワークシートのみのブックと自分で確認した（混在・不明ではチェックしない）</label>
      <label>確認用ブック名（観測値と完全一致）<input value={bookName} onChange={e => setBookName(e.target.value)} /></label>
      <fieldset><legend>対象シートを名前で選択</legend>
        {snapshot?.sheets.map(sheet => <label key={sheet.id}><input type="checkbox" aria-label={`対象シート ${sheet.name}`}
          checked={target.split(/[,\n]/).map(id=>id.trim()).includes(sheet.id)}
          onChange={e=>{const ids=target.split(/[,\n]/).map(id=>id.trim()).filter(Boolean);setTarget((e.target.checked ? [...ids,sheet.id] : ids.filter(id=>id!==sheet.id)).join(","));setPlan(undefined);}} />{sheet.name}（{sheet.visibility}）</label>)}
      </fieldset>
      {kind === "move" && <label>移動先を名前で選択<select value={boundary} onChange={e=>{setBoundary(e.target.value);setPlan(undefined);}}>
        {snapshot?.sheets.map((sheet,index)=><option key={sheet.id} value={index}>{index===0 ? "先頭：" : ""}{sheet.name} の前</option>)}
        <option value={snapshot?.sheets.length ?? 0}>末尾</option>
      </select></label>}
      <label>対象ID（移動はカンマ区切りで複数指定）<textarea value={target} onChange={e => {setTarget(e.target.value);setPlan(undefined);}} /></label>
      <label>操作<select value={kind} onChange={e => {setKind(e.target.value);setPlan(undefined);}}>
        <option>Visible</option><option>Hidden</option><option>VeryHidden</option>
        <option value="copy">複製（直接API・元の表示状態を変えない）</option><option value="delete">直接削除</option>
        <option value="deleteVeryHidden">VeryHidden削除実験（復元保証なし・本番採用不可）</option><option value="move">単一・複数移動</option>
      </select></label>
      <label>移動先境界（元一覧の0～件数、0=先頭、件数=末尾）<input type="number" min="0" value={boundary} onChange={e => {setBoundary(e.target.value);setPlan(undefined);}} /></label>
      <label>合成失敗のステップ番号（0=無効、1から）<input type="number" min="0" value={failureStep} onChange={e => {setFailureStep(e.target.value);setPlan(undefined);}} /></label>
      <select aria-label="合成失敗のタイミング" value={failureWhen} onChange={e => {setFailureWhen(e.target.value as "before"|"after");setPlan(undefined);}}><option value="before">書き込み前</option><option value="after">書き込み後・確認前</option></select>
      <button onClick={prepare} disabled={!enabled || !snapshot}>計画を作成（変更なし）</button>
      <button onClick={requestConfirmation} disabled={!canRun}>確認して実行</button>
    </fieldset>
    {confirmation && <ConfirmationDialog plan={confirmation} onCancel={cancelConfirmation} onExecute={confirmExecution} />}
    {fault && <p role="alert">{fault}</p>}
    <details open><summary>計画（再取得しても固定・不一致は中断）</summary><pre>{JSON.stringify(plan,null,2)}</pre></details>
    <details open><summary>操作前後・結果・エラー（手動コピーして記録）</summary><pre>{JSON.stringify(output,null,2)}</pre></details>
    <details><summary>イベント発火履歴</summary><pre>{JSON.stringify(events,null,2)}</pre></details>
  </section>;
}
