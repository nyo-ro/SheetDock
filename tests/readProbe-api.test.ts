import { afterEach, expect, it, vi } from "vitest";
import { prepareKnownSheets, ProbeKind, runReadProbe } from "../src/dev/readProbe/api";

function setup() {
  let synced = false;
  const target = { id: "private-id", name: "private-name", isNullObject: false, load: vi.fn() };
  const active = { get id() { if (!synced) throw new Error("ID read before sync"); return "private-active-id"; }, load: vi.fn() };
  const sheets = {
    load: vi.fn(),
    get items() { if (!synced) throw new Error("Items read before sync"); return [{id:"private-id",name:"private-name",position:0}]; },
    getItemOrNullObject: vi.fn(() => target), getActiveWorksheet: vi.fn(() => active)
  };
  const context = { workbook: { worksheets: sheets }, sync: vi.fn(async () => { synced = true; }) };
  const run = vi.fn(async (work: (ctx: Excel.RequestContext) => Promise<unknown>) => work(context as unknown as Excel.RequestContext));
  vi.stubGlobal("Excel", { run });
  return { context, target, active, sheets, run, markSynced: () => { synced = true; } };
}
afterEach(() => { vi.unstubAllGlobals(); vi.useRealTimers(); });

it.each(["A", "B", "C"] as ProbeKind[])("%sはExcel.runとsyncを1回だけ実行し開始・終了を記録", async kind => {
  const { context, run } = setup(); const trace = vi.fn();
  await runReadProbe(kind, "private-id", trace);
  expect(run).toHaveBeenCalledOnce(); expect(context.sync).toHaveBeenCalledOnce();
  expect(trace.mock.calls.map(c => c[0])).toEqual(["api.run.start", "api.sync.start", "api.sync.end", "api.run.end"]);
  for (const call of trace.mock.calls.filter(c => c[0].endsWith(".end"))) expect(call[2]).toBeGreaterThanOrEqual(0);
  expect(JSON.stringify(trace.mock.calls)).not.toContain("private");
});

it("Aは一覧のID・名前・位置だけをloadし、sync後にローカル選択肢へ返す", async () => {
  const { sheets, context } = setup(); const trace = vi.fn();
  const original = context.sync.getMockImplementation()!;
  context.sync.mockImplementation(async () => { expect(sheets.load).toHaveBeenCalledWith("items/id,items/name,items/position"); await original(); });
  expect(await runReadProbe("A", undefined, trace)).toEqual({count:1,choices:[{id:"private-id",name:"private-name",position:0}]});
  expect(sheets.getActiveWorksheet).not.toHaveBeenCalled(); expect(sheets.getItemOrNullObject).not.toHaveBeenCalled();
});

it("Bは既知IDの対象だけを読み、アクティブ・一覧取得を行わない", async () => {
  const { sheets, target, context } = setup();
  const original = context.sync.getMockImplementation()!;
  context.sync.mockImplementation(async () => { expect(target.load).toHaveBeenCalledWith("id,name"); await original(); });
  expect(await runReadProbe("B", "private-id", vi.fn())).toEqual({count:1});
  expect(sheets.getItemOrNullObject).toHaveBeenCalledWith("private-id"); expect(sheets.load).not.toHaveBeenCalled(); expect(sheets.getActiveWorksheet).not.toHaveBeenCalled();
});

it("CはアクティブIDだけをloadし、IDを結果や診断へ返さない", async () => {
  const { sheets, active, context } = setup(); const trace = vi.fn();
  const original = context.sync.getMockImplementation()!;
  context.sync.mockImplementation(async () => { expect(active.load).toHaveBeenCalledWith("id"); await original(); });
  expect(await runReadProbe("C", undefined, trace)).toEqual({count:1});
  expect(sheets.getActiveWorksheet).toHaveBeenCalledOnce(); expect(sheets.load).not.toHaveBeenCalled(); expect(sheets.getItemOrNullObject).not.toHaveBeenCalled();
  expect(JSON.stringify(trace.mock.calls)).not.toContain("private");
});

it.each([undefined, "", "  "])("Bの対象未指定はExcel API開始前に拒否: %s", async id => {
  const { run } = setup(); const trace = vi.fn();
  await expect(runReadProbe("B", id, trace)).rejects.toThrow("対象シート");
  expect(run).not.toHaveBeenCalled(); expect(trace).not.toHaveBeenCalled();
});

it("Bの対象消失をsync後に検出し別APIへ退避しない", async () => {
  const { target, sheets, context, run } = setup(); target.isNullObject = true;
  await expect(runReadProbe("B", "private-id", vi.fn())).rejects.toThrow("見つかりません");
  expect(context.sync).toHaveBeenCalledOnce(); expect(run).toHaveBeenCalledOnce(); expect(sheets.getActiveWorksheet).not.toHaveBeenCalled();
});

it("sync失敗を場所とともにtraceへ渡し再試行しない", async () => {
  const { context, run } = setup(); const trace = vi.fn();
  const failure = Object.assign(new Error("private error text"), {code:"InvalidOperation",debugInfo:{errorLocation:"Worksheet.id"}});
  context.sync.mockRejectedValueOnce(failure);
  await expect(runReadProbe("C", undefined, trace)).rejects.toBe(failure);
  expect(run).toHaveBeenCalledOnce(); expect(context.sync).toHaveBeenCalledOnce();
  expect(trace.mock.calls.map(c => c[0])).toEqual(["api.run.start","api.sync.start","api.sync.error","api.sync.end","api.run.error","api.run.end"]);
  expect(trace).toHaveBeenCalledWith("api.sync.error", failure); expect(trace).toHaveBeenCalledWith("api.run.error", failure);
  // エラーは呼出側でcode/locationだけへ変換する。stageへ本文・ID・名前を埋め込まない。
  expect(trace.mock.calls.every(c => !c[0].includes("private"))).toBe(true);
});

it("syncが返らない間は終了ログ・追加API・自動再試行を行わない", async () => {
  vi.useFakeTimers(); const { context, run, markSynced } = setup(); const trace = vi.fn(); let finish!: () => void;
  context.sync.mockImplementationOnce(() => new Promise<void>(resolve => { finish = () => { markSynced(); resolve(); }; }));
  const pending = runReadProbe("C", undefined, trace);
  await vi.advanceTimersByTimeAsync(160000);
  expect(trace.mock.calls.map(c => c[0])).toEqual(["api.run.start","api.sync.start"]);
  expect(run).toHaveBeenCalledOnce(); expect(context.sync).toHaveBeenCalledOnce();
  finish(); await pending; expect(trace.mock.calls[trace.mock.calls.length - 1]?.[0]).toBe("api.run.end");
});

it("Bの準備もactiveを使わず一度の一覧取得だけで選択肢を返す", async () => {
  const { run, sheets, context } = setup(); const trace = vi.fn();
  expect(await prepareKnownSheets(trace)).toEqual([{id:"private-id",name:"private-name",position:0}]);
  expect(run).toHaveBeenCalledOnce(); expect(context.sync).toHaveBeenCalledOnce(); expect(sheets.getActiveWorksheet).not.toHaveBeenCalled();
});

it("未知の実験種別はアクティブ取得へ流さずAPI開始前に拒否", async () => {
  const { run } = setup(); const trace = vi.fn();
  await expect(runReadProbe("invalid" as ProbeKind, undefined, trace)).rejects.toMatchObject({code:"ProbeInvalidKind"});
  expect(run).not.toHaveBeenCalled(); expect(trace).not.toHaveBeenCalled();
});

it("BはID/name両対応APIが別対象を返しても同じ既知IDとみなさない", async () => {
  const { target, run, context } = setup(); target.id = "different-id";
  await expect(runReadProbe("B", "private-id", vi.fn())).rejects.toMatchObject({code:"ProbeTargetMismatch"});
  expect(run).toHaveBeenCalledOnce(); expect(context.sync).toHaveBeenCalledOnce();
});
