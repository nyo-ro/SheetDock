import { afterEach, expect, it, vi } from "vitest";
import { readWorkbookState } from "../src/excel/workbookState";
import { reorderWorksheet } from "../src/excel/reorderWorksheet";
import { workbook } from "./production-fixture";
// Officeホストだけをモック。表示取得・共通キュー・移動前後の検証は実コードを通す。
function host(total: number | null = 3) {
  const state = workbook(); let totalValue = total;
  const totalRead = vi.fn(); const positionWrite = vi.fn();
  let gate: (() => Promise<void>) | undefined;
  const syncCalls = vi.fn();
  const run = vi.fn(async (work: (context: Excel.RequestContext) => Promise<unknown>) => {
    let pendingTotal = false;
    const context = { workbook: {
      name: state.name, readOnly: false, load: () => {}, protection: { protected: false, load: () => {} },
      worksheets: {
        load: () => {}, get items() { return state.sheets.map(s => ({...s, protection:{protected:s.protected}})); },
        getActiveWorksheet: () => ({ id: state.activeId, load: () => {} }),
        getItem: (id: string) => ({ set position(position: number) {
          positionWrite(id, position); const index = state.sheets.findIndex(s => s.id === id);
          const [sheet] = state.sheets.splice(index,1); state.sheets.splice(position,0,sheet); state.sheets.forEach((s,i)=>{s.position=i;});
        } })
      },
      functions: { sheets: () => { totalRead(); pendingTotal = true; return { get value() { return totalValue; }, error:"", load: () => {} }; } }
    }, sync: async () => { syncCalls(); if(pendingTotal) { pendingTotal=false; const waiting=gate;gate=undefined;if(waiting)await waiting(); } } };
    return work(context as unknown as Excel.RequestContext);
  });
  vi.stubGlobal("Office", { context: { document: { url:"https://example.test/book" } } }); vi.stubGlobal("Excel",{run});
  return {state,run,totalRead,positionWrite,syncCalls,waitForNextTotal:(wait:()=>Promise<void>)=>{gate=wait;},setTotal:(value:number|null)=>{totalValue=value;}};
}
afterEach(()=>vi.unstubAllGlobals());
it("表示は1sync・総数なし、移動は総数を前後2回確認してから結果を返す",async()=>{
  const h=host();const baseline=await readWorkbookState();expect(h.syncCalls).toHaveBeenCalledOnce();expect(h.totalRead).not.toHaveBeenCalled();expect(baseline.totalSheets).toBeNull();
  const result=await reorderWorksheet(baseline,"b",0);
  expect(result.error).toBeUndefined();expect(result.state?.sheets.map(s=>s.id)).toEqual(["b","a","c"]);expect(result.state?.activeId).toBe("a");
  expect(h.totalRead).toHaveBeenCalledTimes(2);expect(h.positionWrite).toHaveBeenCalledExactlyOnceWith("b",0);expect(h.syncCalls).toHaveBeenCalledTimes(6);
});
it.each([null,4])("一覧は返せても移動直前の総数%sなら書き込まない",async total=>{
  const h=host(total);const baseline=await readWorkbookState();expect(baseline.sheets).toHaveLength(3);
  const result=await reorderWorksheet(baseline,"b",0);expect(result.error).toContain("混在");expect(h.positionWrite).not.toHaveBeenCalled();expect(h.state.sheets.map(s=>s.id)).toEqual(["a","b","c"]);
});
it("安全確認のSHEETSが未応答なら書き込みも後続表示取得も開始しない",async()=>{
  const h=host();const baseline=await readWorkbookState();let finish!:()=>void;h.waitForNextTotal(()=>new Promise<void>(resolve=>{finish=resolve;}));
  const moving=reorderWorksheet(baseline,"b",0);
  for(let i=0;i<30;i++)await Promise.resolve();expect(h.totalRead).toHaveBeenCalledOnce();expect(h.positionWrite).not.toHaveBeenCalled();
  const queued=readWorkbookState();for(let i=0;i<30;i++)await Promise.resolve();expect(h.run).toHaveBeenCalledTimes(2);expect(h.positionWrite).not.toHaveBeenCalled();
  finish();await moving;const refreshed=await queued;expect(h.run).toHaveBeenCalledTimes(3);expect(h.totalRead).toHaveBeenCalledTimes(2);expect(refreshed.totalSheets).toBeNull();expect(refreshed.sheets[0].id).toBe("b");
});
