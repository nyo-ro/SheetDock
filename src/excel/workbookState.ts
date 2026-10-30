import { diagnose, measured } from "../diagnostics";
import { WorkbookState } from "../domain/workbookState";
import { SheetVisibility } from "../domain/sheets";
import { withExcelLock } from "./serial";
// 呼び出し元がAPIロックを所有する。表示と安全確認で共通のメタデータ。
async function loadMetadata(context: Excel.RequestContext, valid:()=>boolean, includeActive:boolean, display:boolean): Promise<WorkbookState> {
  const check=()=>{if(!valid())throw Object.assign(new Error("終了済みの取得処理を中断しました。"),{code:"SheetDockReadCancelled"});};
  check();const book = context.workbook;
  book.load("name,readOnly"); book.protection.load("protected");
  book.worksheets.load("items/id,items/name,items/position,items/visibility,items/protection/protected");
  // 一覧とアクティブIDは依存しないloadなので同一バッチで観測する。
  // このバッチが失敗したらプロキシを使い回さず、呼出側で新Contextの一覧だけに退避する。
  const active=includeActive ? book.worksheets.getActiveWorksheet() : undefined;
  active?.load("id");
  await measured(display ? (includeActive ? "api.sync.displayMetadataActive" : "api.sync.displayMetadataOnly") : "api.sync.metadataActive",()=>context.sync());
  check();
  const state: WorkbookState = { name:book.name, bookKey:JSON.stringify([Office.context.document.url || "",book.name]),
    readOnly:book.readOnly, structureProtected:book.protection.protected, activeId:null,totalSheets:null,
    sheets:book.worksheets.items.map(s=>({id:s.id,name:s.name,position:s.position,visibility:s.visibility as SheetVisibility,protected:s.protection.protected,isActive:false})).sort((a,b)=>a.position-b.position) };
  if(active && state.sheets.some(s=>s.id===active.id))state.activeId=active.id;
  check();
  state.sheets.forEach(s=>{s.isActive=s.id===state.activeId;});
  return state;
}
// 通常表示専用。総数は未確認のnullで返し、移動の許可には使わない。
export async function loadDisplayState(context: Excel.RequestContext, valid:()=>boolean=()=>true, includeActive=true): Promise<WorkbookState> {
  const state=await loadMetadata(context,valid,includeActive,true);
  return {...state,readPurpose:"display"};
}
// 変更操作前後の厳密な取得。総数評価を省略しない。関数評価はセルを変更しない。
export async function loadWorkbookState(context: Excel.RequestContext, valid:()=>boolean=()=>true): Promise<WorkbookState> {
  const state=await loadMetadata(context,valid,true,false);
  const check=()=>{if(!valid())throw Object.assign(new Error("終了済みの取得処理を中断しました。"),{code:"SheetDockReadCancelled"});};
  check();
  try {
    const count=context.workbook.functions.sheets();count.load("value,error");await measured("api.sync.total",()=>context.sync());
    if (!count.error && typeof count.value === "number" && Number.isInteger(count.value) && count.value > 0) state.totalSheets=count.value;
  } catch { /* SHEETS未対応・評価失敗を通常シートのみと推定しない。 */ }
  check();
  state.sheets.forEach(s=>{s.isActive=s.id===state.activeId;});
  return state;
}
export const readWorkbookState = (valid:()=>boolean=()=>true): Promise<WorkbookState> => withExcelLock(async()=>{
  if(!valid())throw Object.assign(new Error("終了済みの取得処理を中断しました。"),{code:"SheetDockReadCancelled"});
  try {
    return await measured("api.run.read",()=>Excel.run(context=>loadDisplayState(context,valid)));
  } catch(error) {
    if(!valid() || (error as {code?:string})?.code==="SheetDockReadCancelled")throw error;
    // 失敗と確定した後だけ1回。未応答のAPIと重ねたりアクティブ取得を繰り返したりしない。
    diagnose("normal.read.metadataFallback",error);
    const state=await measured("api.run.readMetadataOnly",()=>Excel.run(context=>loadDisplayState(context,valid,false)));
    state.readIssue="アクティブ情報を含む取得に失敗したため、一覧のみ再取得しました。アクティブ表示と移動は無効です。手動更新で再確認してください。";
    return state;
  }
});
