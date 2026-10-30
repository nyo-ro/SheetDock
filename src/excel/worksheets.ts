import { withExcelLock } from "./serial";
import { SheetInfo, SheetVisibility, sortSheets } from "../domain/sheets";

export async function readWorksheets(): Promise<SheetInfo[]> {
  return withExcelLock(() => Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const active = sheets.getActiveWorksheet();
    sheets.load("items/id,items/name,items/position,items/visibility");
    active.load("id");
    await context.sync();
    return sortSheets(sheets.items.map(sheet => ({
      id: sheet.id, name: sheet.name, position: sheet.position,
      visibility: sheet.visibility as SheetVisibility, isActive: sheet.id === active.id
    })));
  }));
}

export async function activateWorksheet(id: string, valid:()=>boolean=()=>true): Promise<void> {
  await withExcelLock(() => Excel.run(async context => {
    if(!valid())throw new Error("終了・中断した操作です。");
    const sheet = context.workbook.worksheets.getItemOrNullObject(id);
    sheet.load("id,visibility");
    await context.sync();
    if (sheet.isNullObject) throw new Error("対象のシートが見つかりません。手動更新してください。");
    if (sheet.visibility !== "Visible") {
      throw new Error("非表示シートには切り替えられません。表示状態は変更していません。");
    }
    if(!valid())throw new Error("終了・中断した操作です。");
    sheet.activate();
    await context.sync();
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("id");
    await context.sync();
    if (active.id !== id) throw new Error("シートの切り替えを確認できませんでした。手動更新してください。");
  }));
}
