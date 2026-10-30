import { moveRestriction, dragRestriction } from "../src/domain/workbookState";
import { afterEach, expect, it, vi } from "vitest";
import { withExcelLock } from "../src/excel/serial";
import { loadWorkbookState, loadDisplayState, readWorkbookState } from "../src/excel/workbookState";
function context(value:unknown=3,error=""){
 const count={value,error,load:vi.fn()};
 const ctx={workbook:{name:"Book.xlsx",readOnly:false,load:vi.fn(),protection:{protected:false,load:vi.fn()},functions:{sheets:vi.fn(()=>count)},worksheets:{load:vi.fn(),items:[0,1,2].map(position=>({id:String(position),name:String(position),position,visibility:"Visible",protection:{protected:false}})),getActiveWorksheet:()=>({id:"0",load:vi.fn()})}},sync:vi.fn(async()=>{})};
 vi.stubGlobal("Office",{context:{document:{url:"https://example.test/book"}}});return ctx;
}
afterEach(()=>vi.unstubAllGlobals());
it("SHEETSを引数なしで評価し、通常シート一覧と独立した総数を取得",async()=>{const ctx=context(4);const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(state.totalSheets).toBe(4);expect(state.sheets).toHaveLength(3);expect(ctx.workbook.functions.sheets).toHaveBeenCalledWith();expect(state.sheets[0].isActive).toBe(true);});
it.each([[null,""],[3,"#VALUE!"],["3",""],[3.5,""],[0,""]])("無効な総数は不明として一覧は維持",async(value,error)=>{const state=await loadWorkbookState(context(value,error as string) as unknown as Excel.RequestContext);expect(state.totalSheets).toBeNull();expect(state.sheets).toHaveLength(3);});
it("SHEETSのAPI例外でも一覧は維持し総数を推測しない",async()=>{const ctx=context();ctx.workbook.functions.sheets.mockImplementation(()=>{throw new Error("unsupported");});const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(state.totalSheets).toBeNull();expect(state.activeId).toBe("0");});

it("Excelロック待ち中に破棄された取得はExcel.runを開始しない",async()=>{
 let finish!:()=>void;const blocker=withExcelLock(()=>new Promise<void>(r=>{finish=r;}));await Promise.resolve();
 const run=vi.fn();vi.stubGlobal("Excel",{run});let valid=true;const read=readWorkbookState(()=>valid);const rejected=expect(read).rejects.toThrow("終了済み");valid=false;finish();await blocker;await rejected;expect(run).not.toHaveBeenCalled();
});
it("メタデータ取得中の破棄後は追加APIを送信しない",async()=>{
 const ctx=context();let valid=true;ctx.sync.mockImplementation(async()=>{valid=false;});
 await expect(loadWorkbookState(ctx as unknown as Excel.RequestContext,()=>valid)).rejects.toThrow("終了済み");expect(ctx.sync).toHaveBeenCalledOnce();expect(ctx.workbook.functions.sheets).not.toHaveBeenCalled();
});

it("一覧とアクティブIDを初回sync前にloadし厳密な安全取得は2sync",async()=>{
 const ctx=context();const active={id:"1",load:vi.fn()};ctx.workbook.worksheets.getActiveWorksheet=()=>active;
 ctx.sync.mockImplementation(async()=>{expect(ctx.workbook.worksheets.load).toHaveBeenCalledOnce();expect(active.load).toHaveBeenCalledWith("id");});
 const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(ctx.sync).toHaveBeenCalledTimes(2);expect(state.activeId).toBe("1");expect(state.sheets[1].isActive).toBe(true);
});
it("アクティブIDが一覧外なら推測せず移動を無効化",async()=>{
 const ctx=context();ctx.workbook.worksheets.getActiveWorksheet=()=>({id:"chart-id",load:vi.fn()});const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(state.activeId).toBeNull();expect(state.sheets.every(s=>!s.isActive)).toBe(true);
});
it("統合バッチ失敗後だけ新Contextで一覧を1回再取得する",async()=>{
 const first=context(),fallback=context();first.sync.mockRejectedValueOnce(new Error("ItemNotFound mock"));const active=vi.fn(()=>({id:"0",load:vi.fn()}));first.workbook.worksheets.getActiveWorksheet=active;fallback.workbook.worksheets.getActiveWorksheet=vi.fn(()=>{throw new Error("must not read active");});
 let index=0;const contexts=[first,fallback];const run=vi.fn(async(fn:(ctx:Excel.RequestContext)=>Promise<unknown>)=>fn(contexts[index++] as unknown as Excel.RequestContext));vi.stubGlobal("Excel",{run});
 const state=await readWorkbookState();expect(run).toHaveBeenCalledTimes(2);expect(active).toHaveBeenCalledOnce();expect(fallback.workbook.worksheets.getActiveWorksheet).not.toHaveBeenCalled();expect(state.sheets).toHaveLength(3);expect(state.activeId).toBeNull();expect(state.readIssue).toContain("一覧のみ");expect(moveRestriction(state)).not.toBeNull();expect(dragRestriction(state)).toContain("アクティブ");expect(state.totalSheets).toBeNull();expect(fallback.sync).toHaveBeenCalledOnce();expect(fallback.workbook.functions.sheets).not.toHaveBeenCalled();
});
it("一覧退避も失敗ならそれ以上再試行しない",async()=>{
 const run=vi.fn(async()=>{throw new Error("offline");});vi.stubGlobal("Excel",{run});await expect(readWorkbookState()).rejects.toThrow("offline");expect(run).toHaveBeenCalledTimes(2);
});
it("破棄された統合バッチの失敗では一覧退避を送らない",async()=>{
 let valid=true;const run=vi.fn(async()=>{valid=false;throw new Error("late failure");});vi.stubGlobal("Excel",{run});await expect(readWorkbookState(()=>valid)).rejects.toThrow("late failure");expect(run).toHaveBeenCalledOnce();
});
it("統合バッチ未応答中は一覧退避も総数取得も開始しない",async()=>{
 const ctx=context();let finish!:()=>void;ctx.sync.mockImplementationOnce(()=>new Promise<void>(r=>{finish=r;}));const run=vi.fn(async(fn:(ctx:Excel.RequestContext)=>Promise<unknown>)=>fn(ctx as unknown as Excel.RequestContext));vi.stubGlobal("Excel",{run});const pending=readWorkbookState();
 for(let i=0;i<10;i++)await Promise.resolve();expect(run).toHaveBeenCalledOnce();expect(ctx.workbook.functions.sheets).not.toHaveBeenCalled();finish();await pending;expect(run).toHaveBeenCalledOnce();
});
it("総数同期完了時にも期限切れなら状態を返さない",async()=>{
 const ctx=context();let valid=true,count=0;ctx.sync.mockImplementation(async()=>{if(++count===2)valid=false;});await expect(loadWorkbookState(ctx as unknown as Excel.RequestContext,()=>valid)).rejects.toThrow("終了済み");
});

it("通常表示は一覧・active・保護を1syncで取得しSHEETSを一切評価しない",async()=>{
 const ctx=context();ctx.workbook.functions.sheets.mockImplementation(()=>{throw new Error("display must not evaluate SHEETS");});
 const active={id:"1",load:vi.fn()};ctx.workbook.worksheets.getActiveWorksheet=()=>active;
 ctx.sync.mockImplementation(async()=>{expect(ctx.workbook.load).toHaveBeenCalledWith("name,readOnly");expect(ctx.workbook.protection.load).toHaveBeenCalledWith("protected");expect(ctx.workbook.worksheets.load).toHaveBeenCalled();expect(active.load).toHaveBeenCalledWith("id");});
 const run=vi.fn(async(fn:(c:Excel.RequestContext)=>Promise<unknown>)=>fn(ctx as unknown as Excel.RequestContext));vi.stubGlobal("Excel",{run});
 const state=await readWorkbookState();expect(run).toHaveBeenCalledOnce();expect(ctx.sync).toHaveBeenCalledOnce();expect(ctx.workbook.functions.sheets).not.toHaveBeenCalled();
 expect(state.sheets).toHaveLength(3);expect(state.activeId).toBe("1");expect(state.sheets[1].isActive).toBe(true);expect(state.totalSheets).toBeNull();expect(state.readPurpose).toBe("display");
 expect(dragRestriction(state)).toBeNull();expect(moveRestriction(state)).not.toBeNull();
});
it("表示取得中に世代が無効になれば後続評価も退避もせず結果を破棄",async()=>{
 const ctx=context();let valid=true;ctx.sync.mockImplementation(async()=>{valid=false;});
 const run=vi.fn(async(fn:(c:Excel.RequestContext)=>Promise<unknown>)=>fn(ctx as unknown as Excel.RequestContext));vi.stubGlobal("Excel",{run});
 await expect(readWorkbookState(()=>valid)).rejects.toMatchObject({code:"SheetDockReadCancelled"});expect(run).toHaveBeenCalledOnce();expect(ctx.sync).toHaveBeenCalledOnce();expect(ctx.workbook.functions.sheets).not.toHaveBeenCalled();
});
it("表示中に未対応のactive IDを得ても総数を推測せずドラッグを禁止",async()=>{
 const ctx=context();ctx.workbook.worksheets.getActiveWorksheet=()=>({id:"chart",load:vi.fn()});
 const state=await loadDisplayState(ctx as unknown as Excel.RequestContext);expect(state.totalSheets).toBeNull();expect(state.activeId).toBeNull();expect(dragRestriction(state)).not.toBeNull();expect(ctx.sync).toHaveBeenCalledOnce();
});
