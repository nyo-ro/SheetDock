import { useEffect, useRef, useState } from "react";
import { subscribeSnapshot } from "./snapshotBus";
import { ensureExcelReady, errorMessage } from "../excel/environment";
import { activateWorksheet } from "../excel/worksheets";
import { readWorkbookState } from "../excel/workbookState";
import { reorderWorksheet } from "../excel/reorderWorksheet";
import { worksheetEvents } from "../excel/worksheetEvents";
import { WorkbookState, moveRestriction } from "../domain/workbookState";
import { CoalescedRefresh } from "../sync/coalescedRefresh";

export function useWorksheets() {
  const [snapshot,setSnapshot]=useState<WorkbookState>();
  const [query,setQuery]=useState("");
  const [ready,setReady]=useState(false);
  const [busy,setBusy]=useState(true);
  const [error,setError]=useState<string|null>(null);
  const [readError,setReadError]=useState<string|null>(null);
  const [syncError,setSyncError]=useState<string|null>(null);
  const mounted=useRef(false), locked=useRef(false), epoch=useRef(0);
  const latest=useRef<WorkbookState>();
  const refreshRef=useRef<()=>Promise<void>>(async()=>{});
  const refreshQueue=useRef<CoalescedRefresh>();
  const apply=(next?:WorkbookState)=>{latest.current=next;if(mounted.current)setSnapshot(next);};
  useEffect(()=>{
    mounted.current=true;let stopped=false;let release:(()=>Promise<void>)|undefined;
    let binding:Promise<void>|undefined;let rebind=false;
    async function connect(){
      if(release)return;
      if(!binding)binding=(async()=>{
        const dispose=await worksheetEvents.subscribe(event=>{
          if(stopped)return;
          if(event.type === "workbookActivated"){epoch.current++;rebind=true;apply(undefined);}
          queue.request();
        });
        if(stopped)await dispose();else release=dispose;
      })().finally(()=>{binding=undefined;});
      await binding;
    }
    async function refresh(){
      await ensureExcelReady();if(stopped)return;
      setReady(true);
      try {
        await connect();if(stopped)return;
        if(rebind){await worksheetEvents.rebind();rebind=false;}
        if(!stopped)setSyncError(null);
      }catch(failure){if(!stopped)setSyncError("自動同期を開始・復旧できません。"+errorMessage(failure));}
      if(stopped)return;
      const requestedEpoch=epoch.current;
      const next=await readWorkbookState();if(stopped)return;
      if(requestedEpoch!==epoch.current){queue.request();return;}
      if(latest.current && latest.current.bookKey!==next.bookKey){
        epoch.current++;rebind=true;
        try{await worksheetEvents.rebind();rebind=false;}catch(failure){setSyncError(errorMessage(failure));}
      }
      if(!stopped){apply(next);setReadError(null);}
    }
    refreshRef.current=refresh;
    const queue=new CoalescedRefresh(refresh,failure=>{
      if(!stopped){apply(undefined);setReadError("自動更新に失敗しました。"+errorMessage(failure));}
    });refreshQueue.current=queue;
    const focus=()=>{if(document.visibilityState!=="hidden")queue.request();};
    window.addEventListener("focus",focus);document.addEventListener("visibilitychange",focus);
    // PoCの結果を直接採用せず、通常UI自身で保護・総シート数も含め再取得する。
    const removeNotice=subscribeSnapshot(()=>queue.request());
    void refresh().catch(failure=>{if(!stopped){apply(undefined);setReadError(errorMessage(failure));}}).finally(()=>{if(!stopped)setBusy(false);});
    return ()=>{
      stopped=true;mounted.current=false;epoch.current++;queue.stop();removeNotice();
      window.removeEventListener("focus",focus);document.removeEventListener("visibilitychange",focus);
      if(release)void release().catch(failure=>console.error("SheetDock event cleanup failed",failure));
    };
  },[]);
  async function run(work:()=>Promise<void>){
    if(!ready || busy || locked.current)return;
    locked.current=true;setBusy(true);setError(null);setReadError(null);
    try{await work();}catch(failure){if(mounted.current){apply(undefined);setError(errorMessage(failure));}}
    finally{locked.current=false;if(mounted.current)setBusy(false);}
  }
  const refresh=()=>run(()=>refreshRef.current());
  const activate=(id:string)=>run(async()=>{const started=epoch.current;await activateWorksheet(id);const next=await readWorkbookState();if(mounted.current && epoch.current===started)apply(next);else refreshQueue.current?.request();});
  const move=(baseline:WorkbookState,id:string,boundary:number,dragEpoch:number)=>run(async()=>{
    if(query)throw new Error("検索をクリアしてから移動してください。");
    const outcome=await reorderWorksheet(baseline,id,boundary,()=>mounted.current && epoch.current===dragEpoch);
    if(mounted.current){if(epoch.current===dragEpoch)apply(outcome.state);else refreshQueue.current?.request();if(outcome.error)setError(outcome.error);}
  });
  return {sheets:snapshot?.sheets??[],snapshot,query,setQuery,ready,busy,error:error ?? readError,syncError,hasSnapshot:!!snapshot,
    moveDisabledReason:query ? "検索中は並び替えできません。検索をクリアしてください。" : syncError || (snapshot ? moveRestriction(snapshot) : "ブック状態を取得中です。"),
    refresh,activate,move,dragEpoch:()=>epoch.current};
}
