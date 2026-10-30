import { afterEach, expect, it, vi } from "vitest";
import { isReadProbeMode, reloadInMode } from "../src/dev/readProbe/mode";
afterEach(() => vi.unstubAllGlobals());
function setup() {
  const items = new Map<string,string>();
  const storage = { getItem: vi.fn((key:string)=>items.get(key) ?? null), setItem:vi.fn((key:string,value:string)=>{items.set(key,value);}), removeItem:vi.fn((key:string)=>{items.delete(key);}) };
  const reload=vi.fn(); vi.stubGlobal("sessionStorage",storage); vi.stubGlobal("window",{location:{reload}}); return {items,storage,reload};
}
it("比較モードを保存してからページ全体を再読み込みする", () => {
  const {items,reload}=setup(); expect(isReadProbeMode()).toBe(false);
  reload.mockImplementation(()=>{expect(items.get("sheetdock.readProbe")).toBe("1");});
  reloadInMode(true); expect(reload).toHaveBeenCalledOnce(); expect(isReadProbeMode()).toBe(true); expect(items.size).toBe(1);
});
it("通常モードへ戻す時もモードだけを解除してページを再起動", () => {
  const {items,reload}=setup(); items.set("sheetdock.readProbe","1"); items.set("unrelated","keep");
  reload.mockImplementation(()=>{expect(items.has("sheetdock.readProbe")).toBe(false);});
  reloadInMode(false); expect(reload).toHaveBeenCalledOnce(); expect(isReadProbeMode()).toBe(false); expect(items.get("unrelated")).toBe("keep");
});
it("ストレージが書けなければ同じ環境でモードを切り替えずエラーを返す", () => {
  const {storage,reload}=setup(); storage.setItem.mockImplementation(()=>{throw new Error("blocked");});
  expect(()=>reloadInMode(true)).toThrow("blocked"); expect(reload).not.toHaveBeenCalled();
});
it("ストレージ読取り不可では比較モードを有効化しない", () => {
  const {storage}=setup(); storage.getItem.mockImplementation(()=>{throw new Error("blocked");}); expect(isReadProbeMode()).toBe(false);
});
