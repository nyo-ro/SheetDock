import { afterEach, describe, expect, it, vi } from "vitest";
import { activateWorksheet, readWorksheets } from "../src/excel/worksheets";
function setup(visibility = "Visible", isNullObject = false) {
  let activeId = "old";
  const sheet = { id: "target", name: "売上", position: 1, visibility, isNullObject,
    load: vi.fn(), activate: vi.fn(() => { activeId = "target"; }) };
  const collection = { items: [sheet, { id: "old", name: "概要", position: 0, visibility: "Visible" }],
    load: vi.fn(), getItemOrNullObject: vi.fn(() => sheet),
    getActiveWorksheet: vi.fn(() => ({ get id() { return activeId; }, load: vi.fn() })) };
  const sync = vi.fn(async () => {});
  vi.stubGlobal("Excel", { run: (body: (context: unknown) => Promise<unknown>) => body({ workbook: { worksheets: collection }, sync }) });
  return { sheet, collection, sync };
}
afterEach(() => vi.unstubAllGlobals());
describe("Excel連携", () => {
  it("ID・名前・表示状態・アクティブ状態を取得して位置順に返す", async () => {
    setup(); const result = await readWorksheets();
    expect(result.map(s => s.id)).toEqual(["old", "target"]);
    expect(result[0].isActive).toBe(true);
    expect(result[1]).toMatchObject({ name: "売上", visibility: "Visible", isActive: false, position: 1 });
  });
  it("取得失敗を呼び出し元に伝える", async () => {
    const { sync } = setup(); sync.mockRejectedValueOnce(new Error("read failed"));
    await expect(readWorksheets()).rejects.toThrow("read failed");
  });
  it("切替直前にIDで再取得し、切替結果を確認する", async () => {
    const { collection, sync, sheet } = setup(); await activateWorksheet("target");
    expect(collection.getItemOrNullObject).toHaveBeenCalledWith("target");
    expect(sheet.activate).toHaveBeenCalledOnce(); expect(sync).toHaveBeenCalledTimes(3);
  });
  it.each(["Hidden", "VeryHidden"])("%sを再表示・切替しない", async visibility => {
    const { sheet } = setup(visibility); await expect(activateWorksheet("target")).rejects.toThrow("非表示");
    expect(sheet.activate).not.toHaveBeenCalled(); expect(sheet.visibility).toBe(visibility);
  });
  it("消失した対象を操作しない", async () => {
    const { sheet } = setup("Visible", true); await expect(activateWorksheet("target")).rejects.toThrow("見つかりません");
    expect(sheet.activate).not.toHaveBeenCalled();
  });
  it("切替のAPI失敗を成功扱いしない", async () => {
    const { sync } = setup(); sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error("activate failed"));
    await expect(activateWorksheet("target")).rejects.toThrow("activate failed");
  });
  it("切替後のID不一致を検出する", async () => {
    const { sheet } = setup(); sheet.activate.mockImplementation(() => {});
    await expect(activateWorksheet("target")).rejects.toThrow("確認できません");
  });
});
