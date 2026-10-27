import { SheetInfo, SheetVisibility } from "../domain/sheets";
export interface ObservedSheet extends SheetInfo { protected: boolean }
export interface BookSnapshot {
  name: string; readOnly: boolean; structureProtected: boolean; sheets: ObservedSheet[];
  activeId: string | null; activeError?: string; capturedAt: string;
}
export type Action = { kind: "visibility"; id: string; visibility: SheetVisibility }
  | { kind: "copy" | "delete"; id: string }
  | { kind: "move"; id: string; position: number; orderAfter: string[] };
export type StepStatus = "success" | "failed" | "unprocessed" | "unknown";
export interface StepResult { action: Action; status: StepStatus; error?: unknown; before?: BookSnapshot; after?: BookSnapshot }
export interface ExperimentResult {
  steps: StepResult[]; before: BookSnapshot; after?: BookSnapshot;
  expectedOrder?: string[]; actualOrder?: string[]; orderMatches?: boolean;
  recovery?: { status: "restored" | "not-needed" | "unknown"; error?: unknown; after?: BookSnapshot };
  safeForProduction: false;
}
export function describeError(error: unknown): object {
  const e = error as { message?: string; code?: string; debugInfo?: unknown } | null;
  return { message: e?.message ?? String(error), code: e?.code, debugInfo: e?.debugInfo };
}
export function fingerprint(book: BookSnapshot): string {
  return JSON.stringify({ name: book.name, readOnly: book.readOnly, structureProtected: book.structureProtected,
    sheets: book.sheets.map(s => [s.id, s.name, s.position, s.visibility, s.protected]) });
}
export function assertFresh(expected: BookSnapshot, actual: BookSnapshot): void {
  if (fingerprint(expected) !== fingerprint(actual)) throw new Error("対象・名前・順序・表示/保護状態が変化しました。計画を作り直してください。");
}
export function assertWritable(book: BookSnapshot, attestedWorksheetsOnly: boolean): void {
  if (!attestedWorksheetsOnly) throw new Error("グラフシート混在・種別不明のブックは観測専用です。");
  if (book.readOnly) throw new Error("読み取り専用ブックです。");
  if (book.structureProtected) throw new Error("ブック構成が保護されています。");
  if (book.activeId === null || !book.sheets.some(s => s.id === book.activeId)) throw new Error("アクティブシートの種類を確認できません。");
  if (book.sheets.some((s, i) => s.position !== i) || new Set(book.sheets.map(s => s.id)).size !== book.sheets.length) {
    throw new Error("位置の欠番またはID重複があります。通常シートのみと確認できません。");
  }
}
export function assertAction(book: BookSnapshot, action: Action): void {
  const target = book.sheets.find(s => s.id === action.id);
  if (!target) throw new Error("対象が消失しました。");
  if (target.protected) throw new Error("シート保護あり。PoCでは保守的に変更を中止します。");
  const removesVisible = action.kind === "delete" || (action.kind === "visibility" && action.visibility !== "Visible");
  if (removesVisible && target.visibility === "Visible" && book.sheets.filter(s => s.visibility === "Visible").length <= 1) {
    throw new Error("最後の表示ワークシートを削除・非表示化できません。");
  }
  if (action.kind === "delete" && target.visibility === "VeryHidden") throw new Error("VeryHiddenはHiddenへ変更してから削除する必要があります。");
  if (action.kind === "move" && (!Number.isInteger(action.position) || action.position < 0 || action.position >= book.sheets.length)) throw new Error("移動先の範囲が不正です。");
}
