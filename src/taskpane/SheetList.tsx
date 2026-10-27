import { DragEvent, useEffect, useRef, useState } from "react";
import { SheetInfo } from "../domain/sheets";
import { WorkbookState, stateSignature } from "../domain/workbookState";
import { groupOrder } from "../domain/movePlan";
const dragType="application/x-sheetdock-worksheet";
interface Props {
  sheets:SheetInfo[]; queryKey?:string; snapshot?:WorkbookState; busy:boolean; disabledReason:string|null;
  activate(id:string):void; move(baseline:WorkbookState,id:string,boundary:number,epoch:number):void; dragEpoch():number;
}
interface DragSession {id:string;baseline:WorkbookState;epoch:number;token:string;displayed:string[];queryKey:string}
export function SheetList({sheets,snapshot,busy,disabledReason,activate,move,dragEpoch,queryKey=""}:Props){
  const session=useRef<DragSession>();const [boundary,setBoundary]=useState<number|null>(null);
  const suppressClick=useRef(false);
  const clear=()=>{session.current=undefined;setBoundary(null);};
  const displayedSignature=JSON.stringify(sheets.map(s=>s.id));
  const currentSignature=snapshot ? stateSignature(snapshot) : "";
  useEffect(()=>{if(session.current && (busy || disabledReason || session.current.queryKey!==queryKey || JSON.stringify(session.current.displayed)!==displayedSignature || stateSignature(session.current.baseline)!==currentSignature))clear();},[busy,disabledReason,currentSignature,displayedSignature,queryKey]);
  useEffect(()=>{const escape=(e:KeyboardEvent)=>{if(e.key==="Escape")clear();};document.addEventListener("keydown",escape);return()=>document.removeEventListener("keydown",escape);},[]);
  function start(e:DragEvent,id:string){
    if(busy || disabledReason || !snapshot || snapshot.sheets.find(s=>s.id===id)?.protected){e.preventDefault();return;}
    const token=`${Date.now()}-${Math.random()}`;
    session.current={id,baseline:structuredClone(snapshot),epoch:dragEpoch(),token,displayed:sheets.map(s=>s.id),queryKey};suppressClick.current=true;
    e.dataTransfer.effectAllowed="move";e.dataTransfer.setData(dragType,token);
  }
  function accept(e:DragEvent):boolean {
    return !!session.current && !busy && !disabledReason && session.current.epoch===dragEpoch() && session.current.queryKey===queryKey && JSON.stringify(session.current.displayed)===displayedSignature && Array.from(e.dataTransfer.types).includes(dragType);
  }
  function over(e:DragEvent,next:number){if(!accept(e))return;e.preventDefault();e.stopPropagation();e.dataTransfer.dropEffect="move";setBoundary(next);}
  function drop(e:DragEvent,next:number){
    const drag=session.current;
    if(!drag || !accept(e) || e.dataTransfer.getData(dragType)!==drag.token){clear();return;}
    e.preventDefault();e.stopPropagation();clear();
    const full=drag.baseline.sheets.map(s=>s.id);const desired=groupOrder(full,drag.displayed,[drag.id],next);
    // 表示境界を元ブックの境界へ変換。検索末尾は最後の表示シートの直後。
    const fullBoundary=next<drag.displayed.length ? full.indexOf(drag.displayed[next]) : full.indexOf(drag.displayed[drag.displayed.length-1])+1;
    if(desired.some((id,i)=>id!==full[i]))move(drag.baseline,drag.id,fullBoundary,drag.epoch);
  }
  const rowBoundary=(e:DragEvent,index:number)=>{const rect=e.currentTarget.getBoundingClientRect();return index+(e.clientY>=rect.top+rect.height/2 ? 1:0);};
  return <ul className="sheets" aria-label="ワークシート一覧" onDragEnd={clear}
    onDragOver={e=>{if(e.target===e.currentTarget)over(e,sheets.length);}}
    onDrop={e=>{if(e.target===e.currentTarget)drop(e,sheets.length);}}
    onDragLeave={e=>{if(!e.currentTarget.contains(e.relatedTarget as Node|null))setBoundary(null);}}>
    {sheets.map((sheet,index)=><li key={sheet.id} className={boundary===index ? "drop-before" : ""}
      onDragOver={e=>over(e,rowBoundary(e,index))} onDrop={e=>drop(e,rowBoundary(e,index))}>
      <div className="sheet-row">
        <span className="drag-handle" draggable={!busy && !disabledReason && !snapshot?.sheets.find(s=>s.id===sheet.id)?.protected}
          aria-label={`${sheet.name}をドラッグして移動`} title={disabledReason ?? "ドラッグして移動"}
          onDragStart={e=>start(e,sheet.id)}>⠿</span>
        <button type="button" className="sheet" aria-current={sheet.isActive ? "true" : undefined}
          disabled={busy || sheet.visibility!=="Visible"}
          draggable={!busy && !disabledReason && !snapshot?.sheets.find(s=>s.id===sheet.id)?.protected}
          onDragStart={e=>start(e,sheet.id)} onPointerDown={()=>{suppressClick.current=false;}}
          onClick={e=>{if(e.detail!==0 && suppressClick.current){suppressClick.current=false;return;}activate(sheet.id);}}>
          <span className="sheet-name">{sheet.name}</span>
          <span className="sheet-state">{sheet.isActive ? "● アクティブ" : sheet.visibility==="Hidden" ? "非表示" : sheet.visibility==="VeryHidden" ? "完全非表示" : ""}</span>
        </button>
      </div>
    </li>)}
    <li className={"drop-end"+(boundary===sheets.length ? " drop-before" : "")} aria-label="一覧末尾の移動先"
      onDragOver={e=>over(e,sheets.length)} onDrop={e=>drop(e,sheets.length)} />
  </ul>;
}
