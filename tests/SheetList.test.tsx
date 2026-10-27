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
it("検索中はドラッグを開始できない",()=>{const s=setup("検索中");expect(screen.getByLabelText("Sheet2をドラッグして移動").getAttribute("draggable")).toBe("false");s.start();s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("ドラッグ中に検索条件が変わっても移動しない",()=>{const s=setup();s.start();s.rerender(<SheetList {...s.props} disabledReason="検索中"/>);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("自動同期で順序が変わったら開始中のドラッグを破棄",()=>{const s=setup();s.start();const updated=structuredClone(s.props.snapshot);updated.sheets[1].name="changed";s.rerender(<SheetList {...s.props} snapshot={updated} sheets={updated.sheets}/>);s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("ブック切替世代が変わればドロップを拒否",()=>{const s=setup();s.start();s.changeEpoch();s.event("drop",0,0);expect(s.move).not.toHaveBeenCalled();});
it("移動中はドラッグ・切替を無効化",()=>{const s=setup();s.rerender(<SheetList {...s.props} busy/>);s.start();s.event("drop",0,0);fireEvent.click(screen.getByRole("button",{name:/Sheet2/}));expect(s.move).not.toHaveBeenCalled();expect(s.activate).not.toHaveBeenCalled();});
