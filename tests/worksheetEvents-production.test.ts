import { afterEach, expect, it, vi } from "vitest";
import { eventSets, registerWorksheetEvents } from "../src/excel/worksheetEvents";
afterEach(()=>vi.unstubAllGlobals());
it("本実装では7イベントとブックアクティブ化を登録・同じcontextで解除",async()=>{
 const removes:ReturnType<typeof vi.fn>[]=[];const add=vi.fn(()=>{const remove=vi.fn();removes.push(remove);return {remove};});
 const context={workbook:{worksheets:Object.fromEntries(Object.keys(eventSets).map(key=>[key,{add}])),onActivated:{add}},sync:vi.fn(async()=>{})};
 const run=vi.fn(async(...args:unknown[])=>{if(args.length===2)expect(args[0]).toBe(context);return (args[args.length-1] as (ctx:unknown)=>Promise<unknown>)(context);});
 vi.stubGlobal("Excel",{run});vi.stubGlobal("Office",{context:{requirements:{isSetSupported:()=>true}}});
 const emit=vi.fn();const dispose=await registerWorksheetEvents(emit,true);expect(add).toHaveBeenCalledTimes(8);
 const callback=(add.mock.calls[7] as unknown as [(args:unknown)=>Promise<void>])[0];await callback({});expect(emit.mock.calls[0][0].type).toBe("workbookActivated");
 await dispose();removes.forEach(remove=>expect(remove).toHaveBeenCalledOnce());
});
