import { describe, expect, it } from "vitest";
import { filterSheets, SheetInfo, sortSheets } from "../src/domain/sheets";
const sheets: SheetInfo[] = [
  { id: "b", name: "Budget 2026", position: 2, visibility: "Hidden", isActive: false },
  { id: "a", name: "売上", position: 0, visibility: "Visible", isActive: true },
  { id: "c", name: "budget 2025", position: 1, visibility: "VeryHidden", isActive: false }
];
describe("一覧と検索", () => {
  it("位置順に並べ、元データを変更しない", () => {
    expect(sortSheets(sheets).map(s => s.id)).toEqual(["a", "c", "b"]);
    expect(sheets[0].id).toBe("b");
  });
  it("大文字小文字を区別せず部分一致する", () => expect(filterSheets(sheets, "BUDGET").map(s => s.id)).toEqual(["b", "c"]));
  it("日本語で検索できる", () => expect(filterSheets(sheets, "売")[0].id).toBe("a"));
  it("クリアで全件に戻る", () => expect(filterSheets(sheets, "")).toEqual(sheets));
  it("該当なしは空", () => expect(filterSheets(sheets, "missing")).toEqual([]));
  it("空白も名前の一部として扱う", () => expect(filterSheets(sheets, " 2026")).toEqual([sheets[0]]));
});
