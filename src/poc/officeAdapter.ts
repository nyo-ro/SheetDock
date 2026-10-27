import { BookSnapshot, Action, describeError } from "./model";
import { SheetVisibility } from "../domain/sheets";
// 呼び出し側がwithExcelLockを保持する。書き込みはこのモジュールに限定。
export async function observeBook(): Promise<BookSnapshot> {
  const snapshot = await Excel.run(async context => {
    const book = context.workbook;
    book.load("name,readOnly"); book.protection.load("protected");
    book.worksheets.load("items/id,items/name,items/position,items/visibility,items/protection/protected");
    await context.sync();
    return { name: book.name, readOnly: book.readOnly, structureProtected: book.protection.protected,
      sheets: book.worksheets.items.map(s => ({ id: s.id, name: s.name, position: s.position,
        visibility: s.visibility as SheetVisibility, protected: s.protection.protected, isActive: false })).sort((a,b) => a.position-b.position),
      activeId: null, capturedAt: new Date().toISOString() } as BookSnapshot;
  });
  // グラフシートがアクティブでも一覧取得の結果を失わないよう別バッチで観測。
  try {
    snapshot.activeId = await Excel.run(async context => {
      const active = context.workbook.worksheets.getActiveWorksheet(); active.load("id"); await context.sync(); return active.id;
    });
    if (!snapshot.sheets.some(s => s.id === snapshot.activeId)) {
      snapshot.activeError = "アクティブIDが通常ワークシート一覧にありません。"; snapshot.activeId = null;
    }
  } catch (error) { snapshot.activeError = JSON.stringify(describeError(error)); }
  try {
    snapshot.totalSheets = await Excel.run(async context => {
      const count = context.workbook.functions.sheets(); count.load("value,error"); await context.sync();
      return !count.error && Number.isInteger(count.value) && count.value > 0 ? count.value : null;
    });
  } catch { snapshot.totalSheets = null; }
  snapshot.sheets.forEach(s => { s.isActive = s.id === snapshot.activeId; });
  return snapshot;
}
export async function applyAction(action: Action): Promise<void> {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(action.id);
    switch (action.kind) {
      case "visibility": sheet.visibility = action.visibility; break;
      case "move": sheet.position = action.position; break;
      case "copy": sheet.copy("After", sheet); break;
      case "delete": sheet.delete(); break;
    }
    await context.sync();
  });
}
