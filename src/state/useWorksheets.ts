import { PendingWatch } from "../sync/pendingWatch";
import { diagnose } from "../diagnostics";
import { useEffect, useRef, useState } from "react";
import { subscribeSnapshot } from "./snapshotBus";
import { ensureExcelReady, errorMessage } from "../excel/environment";
import { activateWorksheet } from "../excel/worksheets";
import { readWorkbookState } from "../excel/workbookState";
import { reorderWorksheet } from "../excel/reorderWorksheet";
import { worksheetEvents } from "../excel/worksheetEvents";
import { WorkbookState, dragRestriction } from "../domain/workbookState";
import { CoalescedRefresh } from "../sync/coalescedRefresh";

export function useWorksheets() {
  const [snapshot,setSnapshot]=useState<WorkbookState>();
  const [query,setQuery]=useState("");
  const [ready,setReady]=useState(false);
  const [busy,setBusy]=useState(true);
  const [error,setError]=useState<string|null>(null);
  const [readError,setReadError]=useState<string|null>(null);
  const [stalled,setStalled]=useState(false);
  const stalledRef=useRef(false);
  const watchRef=useRef<PendingWatch>();
  const revision=useRef(0);
  const [syncError,setSyncError]=useState<string|null>(null);
  const mounted=useRef(false), locked=useRef(false), epoch=useRef(0);
  const latest=useRef<WorkbookState>();
  const refreshRef=useRef<()=>Promise<void>>(async()=>{});
  const refreshQueue=useRef<CoalescedRefresh>();
  const apply=(next?:WorkbookState)=>{latest.current=next;if(mounted.current){setSnapshot(next);if(next?.readIssue)setReadError(next.readIssue);}};
  useEffect(()=>{
    mounted.current=true;diagnose("pane.init");let stopped=false;let release:(()=>Promise<void>)|undefined;
    let binding:Promise<void>|undefined;let rebind=false;let recoveryAvailable=true;
    const watch=new PendingWatch(value=>{stalledRef.current=value;setStalled(value);if(value){apply(undefined);setReadError("Excelの応答に時間がかかり、最新状態を確認できていません。状態が戻らない場合はログ保存後にExcelのブラウザータブを再読み込みしてください。");}},()=>{
      if(recoveryAvailable && !stopped){recoveryAvailable=false;diagnose("normal.refresh.recovery");queue.request();}
    });watchRef.current=watch;
    function request(){revision.current++;recoveryAvailable=true;queue.request();}
    async function connect(){
      if(release)return;
      if(!binding)binding=(async()=>{
        const dispose=await worksheetEvents.subscribe(event=>{
          if(stopped)return;
          if(event.type === "workbookActivated"){epoch.current++;rebind=true;apply(undefined);}
          request();
        });
        if(stopped)await dispose();else release=dispose;
      })().finally(()=>{binding=undefined;});
      await binding;
    }
    async function refresh(){
      await watch.run("normal.refresh",async valid=>{
        await ensureExcelReady();if(stopped || !valid())return;
        setReady(true);
        try {
          await connect();if(stopped || !valid())return;
          if(rebind){const generation=epoch.current;await worksheetEvents.rebind();rebind=epoch.current!==generation;}
          if(!stopped && valid())setSyncError(null);
        }catch(failure){if(!stopped && valid())setSyncError("自動同期を開始・復旧できません。"+errorMessage(failure));}
        if(stopped || !valid())return;
        const requestedEpoch=epoch.current,requestedRevision=revision.current;
        const current=()=>!stopped && valid() && requestedEpoch===epoch.current && requestedRevision===revision.current;
        const discarded=()=>diagnose("normal.refresh.discarded."+(stopped ? "disposed" : !valid() ? "expired" : requestedEpoch!==epoch.current ? "generation" : "revision"));
        let next:WorkbookState;
        try {next=await readWorkbookState(current);}
        catch(failure){if(!current()){discarded();if(stopped || !valid() || (failure as {code?:string})?.code==="SheetDockReadCancelled")return;}throw failure;}
        if(!current()){discarded();return;}
        if(latest.current && latest.current.bookKey!==next.bookKey){
          epoch.current++;rebind=true;const generation=epoch.current;
          try{await worksheetEvents.rebind();rebind=epoch.current!==generation;}catch(failure){if(!stopped)setSyncError(errorMessage(failure));}
          if(stopped || !valid() || epoch.current!==generation || requestedRevision!==revision.current)return;
        }
        if(!stopped && valid()){apply(next);setReadError(next.readIssue ?? null);diagnose("normal.refresh.applied");}
      });
    }
    const queue=new CoalescedRefresh(refresh,failure=>{
      if(!stopped){apply(undefined);setReadError("自動更新に失敗しました。"+errorMessage(failure));}
    },150,"normal.refresh");refreshQueue.current=queue;
    refreshRef.current=()=>{revision.current++;recoveryAvailable=true;return queue.runNow();};
    const focus=()=>{if(document.visibilityState!=="hidden")request();};
    window.addEventListener("focus",focus);document.addEventListener("visibilitychange",focus);
    // PoCの結果を直接採用せず、通常UI自身で表示用情報を再取得する。総数は変更操作前に確認。
    const removeNotice=subscribeSnapshot(request);
    void queue.runNow().finally(()=>{if(!stopped)setBusy(false);});
    return ()=>{
      stopped=true;diagnose("pane.dispose");mounted.current=false;epoch.current++;revision.current++;watch.stop();queue.stop();removeNotice();
      window.removeEventListener("focus",focus);document.removeEventListener("visibilitychange",focus);
      if(release)void release().catch(failure=>console.error("SheetDock event cleanup failed",failure));
    };
  },[]);
  async function run(work:(valid:()=>boolean)=>Promise<void>){
    if(!ready || busy || locked.current || stalledRef.current)return;
    locked.current=true;setBusy(true);setError(null);setReadError(null);
    try{await watchRef.current!.run("normal.operation",work);}catch(failure){if(mounted.current){apply(undefined);setError(errorMessage(failure));}}
    finally{locked.current=false;if(mounted.current)setBusy(false);}
  }
  const refresh=()=>run(()=>refreshRef.current());
  const activate=(id:string)=>run(async valid=>{
    const started=epoch.current;revision.current++;const startedRevision=revision.current;
    await activateWorksheet(id,()=>mounted.current && valid() && epoch.current===started);
    const next=await readWorkbookState(()=>mounted.current && valid());
    if(mounted.current && valid() && epoch.current===started && revision.current===startedRevision)apply(next);else refreshQueue.current?.request();
  });
  const move=(baseline:WorkbookState,id:string,boundary:number,dragEpoch:number)=>run(async valid=>{
    revision.current++;const startedRevision=revision.current;
    const outcome=await reorderWorksheet(baseline,id,boundary,()=>mounted.current && valid() && epoch.current===dragEpoch);
    if(mounted.current){if(valid() && epoch.current===dragEpoch && revision.current===startedRevision)apply(outcome.state);else refreshQueue.current?.request();if(outcome.error)setError(outcome.error);}
  });
  return {sheets:snapshot?.sheets??[],snapshot,query,setQuery,ready:ready && !stalled,busy:busy && !stalled,stalled,operationBlocked:busy || stalled,error:stalled ? "Excelの応答が15秒以上ありません。処理の完了を待っています。返答後に再取得を試みます。戻らない場合は診断ログを保存し、Excelのブラウザータブを再読み込みしてください。" : error ?? readError,syncError,hasSnapshot:!!snapshot,
    moveDisabledReason:stalled ? "Excelの応答待ちのため操作できません。" : syncError || (snapshot ? dragRestriction(snapshot) : readError ? "現在の状態を確認できません。手動更新してください。" : "ブック状態を取得中です。"),
    refresh,activate,move,dragEpoch:()=>epoch.current};
}
