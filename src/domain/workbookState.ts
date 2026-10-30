import { SheetInfo } from "./sheets";
export interface WorkbookState {
  bookKey: string; name: string; readOnly: boolean; structureProtected: boolean;
  sheets: (SheetInfo & { protected: boolean })[];
  activeId: string | null; totalSheets: number | null;
  readIssue?: string;
  // 表示用取得は総数未確認。変更可否は必ず別の厳密な取得で判定する。
  readPurpose?: "display";
}
export function stateSignature(state: WorkbookState): string {
  return JSON.stringify([state.bookKey,state.name,state.readOnly,state.structureProtected,state.totalSheets,
    state.sheets.map(s=>[s.id,s.name,s.position,s.visibility,s.protected])]);
}
// 総数をまだ読んでいない表示スナップショットとは、観測できた項目だけ照合する。
// currentに対するmoveRestriction成功を、書き込み前に必ず別途要求する。
export function matchesMoveBaseline(baseline: WorkbookState, current: WorkbookState): boolean {
  if (baseline.readPurpose !== "display") return stateSignature(baseline) === stateSignature(current);
  if (baseline.totalSheets !== null) return false;
  const observed = (state: WorkbookState) => JSON.stringify([state.bookKey,state.name,state.readOnly,state.structureProtected,
    state.sheets.map(s=>[s.id,s.name,s.position,s.visibility,s.protected])]);
  return observed(baseline) === observed(current);
}
// ドラッグは移動要求の作成のみ。総数未確認でも送信前の厳密な確認へ進める。
export function dragRestriction(state: WorkbookState): string | null {
  return restriction(state, state.readPurpose !== "display" || state.totalSheets !== null);
}
export function moveRestriction(state: WorkbookState): string | null {
  return restriction(state, true);
}
function restriction(state: WorkbookState, requireTotal: boolean): string | null {
  if (state.readOnly) return "読み取り専用のため、シートを移動できません。";
  if (state.structureProtected) return "ブックの構成が保護されているため、シートを移動できません。";
  if ((requireTotal && (!Number.isInteger(state.totalSheets) || state.totalSheets !== state.sheets.length)) ||
      state.sheets.length === 0 || state.sheets.some((s,i)=>s.position !== i) ||
      new Set(state.sheets.map(s=>s.id)).size !== state.sheets.length) {
    return "未対応のシートが混在しているか、ブック構成を確認できないため、移動できません。";
  }
  if (!state.activeId || !state.sheets.some(s=>s.id === state.activeId && s.visibility === "Visible")) return "アクティブシートを確認できないため、移動できません。";
  return null;
}
