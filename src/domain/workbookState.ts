import { SheetInfo } from "./sheets";
export interface WorkbookState {
  bookKey: string; name: string; readOnly: boolean; structureProtected: boolean;
  sheets: (SheetInfo & { protected: boolean })[];
  activeId: string | null; totalSheets: number | null;
}
export function stateSignature(state: WorkbookState): string {
  return JSON.stringify([state.bookKey,state.name,state.readOnly,state.structureProtected,state.totalSheets,
    state.sheets.map(s=>[s.id,s.name,s.position,s.visibility,s.protected])]);
}
export function moveRestriction(state: WorkbookState): string | null {
  if (state.readOnly) return "読み取り専用のため、シートを移動できません。";
  if (state.structureProtected) return "ブックの構成が保護されているため、シートを移動できません。";
  if (!Number.isInteger(state.totalSheets) || state.totalSheets !== state.sheets.length ||
      state.sheets.length === 0 || state.sheets.some((s,i)=>s.position !== i) ||
      new Set(state.sheets.map(s=>s.id)).size !== state.sheets.length) {
    return "未対応のシートが混在しているか、ブック構成を確認できないため、移動できません。";
  }
  if (!state.activeId || !state.sheets.some(s=>s.id === state.activeId && s.visibility === "Visible")) return "アクティブシートを確認できないため、移動できません。";
  return null;
}
