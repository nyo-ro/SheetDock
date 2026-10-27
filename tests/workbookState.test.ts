import { afterEach, expect, it, vi } from "vitest";
import { loadWorkbookState } from "../src/excel/workbookState";
function context(value:unknown=3,error=""){
 const count={value,error,load:vi.fn()};
 const ctx={workbook:{name:"Book.xlsx",readOnly:false,load:vi.fn(),protection:{protected:false,load:vi.fn()},functions:{sheets:vi.fn(()=>count)},worksheets:{load:vi.fn(),items:[0,1,2].map(position=>({id:String(position),name:String(position),position,visibility:"Visible",protection:{protected:false}})),getActiveWorksheet:()=>({id:"0",load:vi.fn()})}},sync:vi.fn(async()=>{})};
 vi.stubGlobal("Office",{context:{document:{url:"https://example.test/book"}}});return ctx;
}
afterEach(()=>vi.unstubAllGlobals());
it("SHEETSを引数なしで評価し、通常シート一覧と独立した総数を取得",async()=>{const ctx=context(4);const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(state.totalSheets).toBe(4);expect(state.sheets).toHaveLength(3);expect(ctx.workbook.functions.sheets).toHaveBeenCalledWith();expect(state.sheets[0].isActive).toBe(true);});
it.each([[null,""],[3,"#VALUE!"],["3",""],[3.5,""],[0,""]])("無効な総数は不明として一覧は維持",async(value,error)=>{const state=await loadWorkbookState(context(value,error as string) as unknown as Excel.RequestContext);expect(state.totalSheets).toBeNull();expect(state.sheets).toHaveLength(3);});
it("SHEETSのAPI例外でも一覧は維持し総数を推測しない",async()=>{const ctx=context();ctx.workbook.functions.sheets.mockImplementation(()=>{throw new Error("unsupported");});const state=await loadWorkbookState(ctx as unknown as Excel.RequestContext);expect(state.totalSheets).toBeNull();expect(state.activeId).toBe("0");});
