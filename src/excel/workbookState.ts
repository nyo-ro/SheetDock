import { WorkbookState } from "../domain/workbookState";
import { SheetVisibility } from "../domain/sheets";
import { withExcelLock } from "./serial";
// 呼び出し元がAPIロックを所有する。関数評価はセルを変更しない。
export async function loadWorkbookState(context: Excel.RequestContext): Promise<WorkbookState> {
  const book = context.workbook;
  book.load("name,readOnly"); book.protection.load("protected");
  book.worksheets.load("items/id,items/name,items/position,items/visibility,items/protection/protected");
  await context.sync();
  const state: WorkbookState = { name:book.name, bookKey:JSON.stringify([Office.context.document.url || "",book.name]),
    readOnly:book.readOnly, structureProtected:book.protection.protected, activeId:null,totalSheets:null,
    sheets:book.worksheets.items.map(s=>({id:s.id,name:s.name,position:s.position,visibility:s.visibility as SheetVisibility,protected:s.protection.protected,isActive:false})).sort((a,b)=>a.position-b.position) };
  try {
    const active=book.worksheets.getActiveWorksheet();active.load("id");await context.sync();
    if (state.sheets.some(s=>s.id===active.id)) state.activeId=active.id;
  } catch { /* 一覧は残すが、移動を許可しない。 */ }
  try {
    const count=book.functions.sheets();count.load("value,error");await context.sync();
    if (!count.error && typeof count.value === "number" && Number.isInteger(count.value) && count.value > 0) state.totalSheets=count.value;
  } catch { /* SHEETS未対応・評価失敗を通常シートのみと推定しない。 */ }
  state.sheets.forEach(s=>{s.isActive=s.id===state.activeId;});
  return state;
}
export const readWorkbookState = (): Promise<WorkbookState> => withExcelLock(()=>Excel.run(loadWorkbookState));
