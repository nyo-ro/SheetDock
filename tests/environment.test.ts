import { afterEach, expect, it, vi } from "vitest";
import { ensureExcelReady } from "../src/excel/environment";
afterEach(() => vi.unstubAllGlobals());
function office(host: string, supported: boolean) {
  vi.stubGlobal("Office", { onReady: async () => ({ host }), HostType: { Excel: "Excel" },
    context: { requirements: { isSetSupported: (name: string, version: string) => supported && name === "ExcelApi" && version === "1.17" } } });
}
it("Office.jsがなければエラー", async () => { vi.stubGlobal("Office", undefined); await expect(ensureExcelReady()).rejects.toThrow("読み込めません"); });
it("Excel以外を拒否する", async () => { office("Word", true); await expect(ensureExcelReady()).rejects.toThrow("Excelのタスクペイン"); });
it("API未対応を拒否する", async () => { office("Excel", false); await expect(ensureExcelReady()).rejects.toThrow("1.17"); });
it("対応環境は成功する", async () => { office("Excel", true); await expect(ensureExcelReady()).resolves.toBeUndefined(); });
it("Office初期化を無期限に待たない", async () => {
  vi.stubGlobal("Office", { onReady: () => new Promise(() => {}) });
  await expect(ensureExcelReady(5)).rejects.toThrow("タイムアウト");
});
