import { WorkbookState, moveRestriction, stateSignature } from "../domain/workbookState";
import { groupOrder } from "../domain/movePlan";
import { loadWorkbookState } from "./workbookState";
import { withExcelLock } from "./serial";
export interface MoveOutcome { state?: WorkbookState; error?: string }
export async function reorderWorksheet(baseline: WorkbookState, id: string, boundary: number, stillValid: () => boolean = () => true): Promise<MoveOutcome> {
  return withExcelLock(async()=>{
    let sent=false;
    try {
      const state=await Excel.run(async context=>{
        const current=await loadWorkbookState(context);
        const blocked=moveRestriction(current);if(blocked)throw new Error(blocked);
        if(!stillValid() || stateSignature(current)!==stateSignature(baseline)) throw new Error("ドラッグ開始後にブックやシートの状態が変わりました。移動を中止しました。");
        const target=current.sheets.find(s=>s.id===id);
        if(!target)throw new Error("移動対象のシートが見つかりません。");
        if(target.protected)throw new Error("保護されたシートの移動には対応していません。");
        const full=current.sheets.map(s=>s.id);const desired=groupOrder(full,full,[id],boundary);
        if(desired.every((s,i)=>s===full[i]))return current;
        context.workbook.worksheets.getItem(id).position=desired.indexOf(id);sent=true;await context.sync();
        const after=await loadWorkbookState(context);
        const expected:WorkbookState={...current,sheets:desired.map((key,position)=>({...current.sheets.find(s=>s.id===key)!,position}))};
        if(moveRestriction(after) || stateSignature(after)!==stateSignature(expected) || after.activeId!==current.activeId) throw new Error("移動後の状態が期待と一致しません。");
        return after;
      });
      return {state};
    } catch(error) {
      const detail=error instanceof Error ? error.message : "Excelの操作に失敗しました。";
      const message=sent ? `移動結果を確認できません。変更が適用された可能性があります。${detail}` : detail;
      try {return {error:message,state:await Excel.run(loadWorkbookState)};}
      catch {return {error:message+" 現在状態の再取得にも失敗しました。手動更新してください。"};}
    }
  });
}
