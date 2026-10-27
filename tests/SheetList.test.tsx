// @vitest-environment jsdom
import { afterEach, expect, it, vi } from "vitest";
import { cleanup, createEvent, fireEvent, render, screen } from "@testing-library/react";
import { SheetList } from "../src/taskpane/SheetList";
import { workbook } from "./production-fixture";
afterEach(cleanup);
function transfer(){const data=new Map<string,string>();return {effectAllowed:"",dropEffect:"",get types(){return [...data.keys()];},setData:(key:string,value:string)=>data.set(key,value),getData:(key:string)=>data.get(key)??""};}
function setup(reason:string|null=null){
 const snapshot=workbook(),move=vi.fn(),activate=vi.fn();let epoch=0;
 const props={sheets:snapshot.sheets,snapshot,busy:false,disabledReason:reason,move,activate,dragEpoch:()=>epoch};
 const result=render(<SheetList {...props}/>);
 const list=screen.getByRole("list");const rows=list.querySelectorAll(":scope > li");
 rows.forEach((row,index)=>vi.spyOn(row,"getBoundingClientRect").mockReturnValue({top:index*40,height:40} as DOMRect));
 const dt=transfer();
 const start=(name="Sheet2")=>fireEvent.dragStart(screen.getByLabelText(`${name}をドラッグして移動`),{dataTransfer:dt});
 const event=(type:"dragOver"|"drop",row:number,y:number)=>{const e=createEvent[type](rows[row],{dataTransfer:dt});Object.defineProperty(e,"clientY",{value:y});fireEvent(rows[row],e);};
 return {...result,props,rows,list,dt,start,event,move,activate,changeEpoch:()=>epoch++};
}
it.each([["Sheet2",0,0,0],["Sheet1",3,140,3],["Sheet3",1,40,1]] as const)("%sを境界へドロップし開始時スナップショットを渡す",(name,row,y,boundary)=>{
 const s=setup();s.start(name);s.event("dragOver",row,y);expect(s.rows[row].className).toContain("drop-before");s.event("drop",row,y);
 expect(s.move).toHaveBeenCalledWith(s.props.snapshot,name==="Sheet1"?"a":name==="Sheet2"?"b":"c",boundary,0);expect(s.activate).not.toHaveBeenCalled();
});
it("同じ位置ではAPI層へ進まない",()=>{const s=setup();s.start();s.event("drop",1,40);expect(s.move).not.toHaveBeenCalled();});
it("一覧外ドロップとdragend後は変更しない",()=>{const s=setup();s.start();s.event("dragOver",0,0);fireEvent.drop(document.body,{dataTransfer:s.dt});fireEvent.dragEnd(s.list);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("Escapeで中断したドラッグは変更しない",()=>{const s=setup();s.start();s.event("dragOver",0,0);fireEvent.keyDown(document,{key:"Escape"});s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("外部からのドロップを拒否",()=>{const s=setup();s.dt.setData("application/x-sheetdock-worksheet","foreign");s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("連続dropは一度だけ",()=>{const s=setup();s.start();s.event("drop",0,0);s.event("drop",0,0);expect(s.move).toHaveBeenCalledTimes(1);});
it("安全性の無効理由があればドラッグを開始できない",()=>{const s=setup("構成不明");expect(screen.getByLabelText("Sheet2をドラッグして移動").getAttribute("draggable")).toBe("false");s.start();s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("ドラッグ中に検索条件が変わっても移動しない",()=>{const s=setup();s.start();s.rerender(<SheetList {...s.props} queryKey="new query"/>);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("自動同期で順序が変わったら開始中のドラッグを破棄",()=>{const s=setup();s.start();const updated=structuredClone(s.props.snapshot);updated.sheets[1].name="changed";s.rerender(<SheetList {...s.props} snapshot={updated} sheets={updated.sheets}/>);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("ブック切替世代が変わればドロップを拒否",()=>{const s=setup();s.start();s.changeEpoch();s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("移動中はドラッグ・切替を無効化",()=>{const s=setup();s.rerender(<SheetList {...s.props} busy/>);s.start();s.event("drop",0,0);fireEvent.click(screen.getByRole("button",{name:/Sheet2/}));expect(s.move).not.toHaveBeenCalled();expect(s.activate).not.toHaveBeenCalled();});

it.each([["Sheet3",0,0,0],["Sheet1",1,60,3]] as const)("絞り込み中%sの移動先を全体境界へ変換",(name,row,y,gap)=>{
 const s=setup();s.rerender(<SheetList {...s.props} sheets={s.props.sheets.filter(sheet=>sheet.id!=="b")} queryKey="filtered"/>);
 const rows=s.list.querySelectorAll(":scope > li");rows.forEach((r,i)=>vi.spyOn(r,"getBoundingClientRect").mockReturnValue({top:i*40,height:40} as DOMRect));
 s.start(name);const e=createEvent.drop(rows[row],{dataTransfer:s.dt});Object.defineProperty(e,"clientY",{value:y});fireEvent(rows[row],e);
 expect(s.move).toHaveBeenCalledWith(s.props.snapshot,name==="Sheet1"?"a":"c",gap,0);
});
it("絞り込み中のEscでも移動しない",()=>{const s=setup();s.rerender(<SheetList {...s.props} sheets={[s.props.sheets[0],s.props.sheets[2]]} queryKey="filtered"/>);s.start("Sheet3");fireEvent.keyDown(document,{key:"Escape"});s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("同じ検索文字列でも表示対象が変わればドラッグを破棄",()=>{const s=setup();s.start();s.rerender(<SheetList {...s.props} sheets={s.props.sheets.slice(1)}/>);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});

it.each([["d",0,1],["b",1,3],["b",2,4]] as const)("対象外を前後に含む検索結果の%sを境界%iへ変換",(id,row,gap)=>{
 const snapshot=workbook();snapshot.sheets=["a","b","c","d","e"].map((id,position)=>({...snapshot.sheets[0],id,name:id,position,isActive:id==="a"}));snapshot.totalSheets=5;
 const move=vi.fn();render(<SheetList snapshot={snapshot} sheets={snapshot.sheets.filter(s=>["b","d"].includes(s.id))} queryKey="filter" busy={false} disabledReason={null} move={move} activate={()=>{}} dragEpoch={()=>0}/>);
 const rows=screen.getByRole("list").querySelectorAll(":scope > li");rows.forEach((r,i)=>vi.spyOn(r,"getBoundingClientRect").mockReturnValue({top:i*40,height:40} as DOMRect));
 const dt=transfer();fireEvent.dragStart(screen.getByLabelText(`${id}をドラッグして移動`),{dataTransfer:dt});
 const e=createEvent.drop(rows[row],{dataTransfer:dt});Object.defineProperty(e,"clientY",{value:row*40});fireEvent(rows[row],e);
 expect(move).toHaveBeenCalledWith(snapshot,id,gap,0);
});
