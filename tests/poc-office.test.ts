import { afterEach, expect, it, vi } from "vitest";
import { eventSets, registerWorksheetEvents } from "../src/poc/events";
import { applyAction, observeBook } from "../src/poc/officeAdapter";
afterEach(()=>vi.unstubAllGlobals());
it("コレクション7イベントを1回ずつ登録し同じcontextで解除する",async()=>{
  const handles=Object.keys(eventSets).map(()=>({remove:vi.fn()}));
  const add=Object.fromEntries(Object.keys(eventSets).map((key,i)=>[key,{add:vi.fn(()=>handles[i])}]));
  const context={workbook:{worksheets:add},sync:vi.fn(async()=>{})};
  const run=vi.fn(async(...args:unknown[])=>{if(args.length===2)expect(args[0]).toBe(context);return (args[args.length-1] as (ctx:unknown)=>Promise<unknown>)(context);});
  vi.stubGlobal("Excel",{run});vi.stubGlobal("Office",{context:{requirements:{isSetSupported:()=>true}}});
  const emit=vi.fn();const dispose=await registerWorksheetEvents(emit);
  Object.values(add).forEach(e=>expect(e.add).toHaveBeenCalledOnce());
  // コールバックはAPI取得せず通知だけ。集約処理が後で再取得する。
  const callback=add.onMoved.add.mock.calls[0] as unknown as [(args:unknown)=>Promise<void>];
  await callback[0]({worksheetId:"a"});expect(emit.mock.calls[0][0].type).toBe("onMoved");expect(run).toHaveBeenCalledTimes(1);
  await dispose();handles.forEach(h=>expect(h.remove).toHaveBeenCalledOnce());expect(context.sync).toHaveBeenCalledTimes(2);
});
it("イベント登録sync失敗時も登録済みハンドラの解除を試みる",async()=>{
  const remove=vi.fn();const add=Object.fromEntries(Object.keys(eventSets).map(key=>[key,{add:()=>({remove})}]));
  const context={workbook:{worksheets:add},sync:vi.fn(async()=>{}).mockRejectedValueOnce(new Error("register failure"))};
  vi.stubGlobal("Excel",{run:async(fn:(ctx:unknown)=>Promise<unknown>)=>fn(context)});vi.stubGlobal("Office",{context:{requirements:{isSetSupported:()=>true}}});
  await expect(registerWorksheetEvents(vi.fn())).rejects.toThrow("register failure");expect(remove).toHaveBeenCalledTimes(7);
});
it("グラフシート等でアクティブ取得が失敗しても一覧と位置を保持する",async()=>{
  const collection={load:vi.fn(),items:[{id:"a",name:"A",position:2,visibility:"Visible",protection:{protected:false}}],getActiveWorksheet:()=>{throw new Error("chartsheet active");}};
  const context={workbook:{name:"Mixed",readOnly:false,load:vi.fn(),protection:{load:vi.fn(),protected:false},worksheets:collection},sync:async()=>{}};
  vi.stubGlobal("Excel",{run:async(fn:(ctx:unknown)=>Promise<unknown>)=>fn(context)});
  const result=await observeBook();expect(result.sheets[0].position).toBe(2);expect(result.activeId).toBeNull();expect(result.activeError).toContain("chartsheet active");
});
it("対象IDのAPIにだけ書き込み、複製ではvisibilityを変更しない",async()=>{
  let visibility="VeryHidden";const setVisibility=vi.fn((value:string)=>{visibility=value;});
  const sheet={get visibility(){return visibility;},set visibility(value:string){setVisibility(value);},copy:vi.fn(),delete:vi.fn(),position:0};
  const getItem=vi.fn(()=>sheet);const sync=vi.fn(async()=>{});vi.stubGlobal("Excel",{run:async(fn:(ctx:unknown)=>Promise<unknown>)=>fn({workbook:{worksheets:{getItem}},sync})});
  await applyAction({kind:"copy",id:"target-id"});expect(getItem).toHaveBeenCalledWith("target-id");expect(sheet.copy).toHaveBeenCalledWith("After",sheet);expect(setVisibility).not.toHaveBeenCalled();
  await applyAction({kind:"visibility",id:"target-id",visibility:"Hidden"});expect(setVisibility).toHaveBeenCalledWith("Hidden");expect(sync).toHaveBeenCalledTimes(2);
});
