import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { ProbeKind, ProbeResult, ProbeTrace, SheetChoice } from "../src/dev/readProbe/api";
import { ProbeDependencies, ReadProbeController } from "../src/dev/readProbe/controller";

function deferred<T>() {
  let resolve!: (value: T) => void; let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((ok, bad) => { resolve = ok; reject = bad; });
  return { promise, resolve, reject };
}
const choices: SheetChoice[] = [{ id: "private-id", name: "private-sheet-name", position: 0 }];
async function traced<T>(trace: ProbeTrace, work: () => Promise<T>): Promise<T> {
  trace("api.run.start"); trace("api.sync.start");
  try { return await work(); }
  catch (error) { trace("api.sync.error", error); trace("api.run.error", error); throw error; }
  finally { trace("api.sync.end", undefined, 7); trace("api.run.end", undefined, 8); }
}
function setup() {
  const dependencies = {
    ready: vi.fn(async () => {}),
    prepare: vi.fn(async (trace: ProbeTrace) => traced(trace, async () => choices.map(s => ({ ...s })))),
    run: vi.fn(async (_kind: ProbeKind, _id: string | undefined, trace: ProbeTrace): Promise<ProbeResult> => traced(trace, async () => ({ count: 1 })))
  } satisfies ProbeDependencies;
  const controller = new ReadProbeController(dependencies);
  return { controller, dependencies };
}
beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(new Date("2026-10-30T00:00:00Z")); });
afterEach(() => { vi.useRealTimers(); });

it("作成時は時刻とセッションだけを記録しExcelを取得しない", () => {
  const { controller, dependencies } = setup();
  expect(controller.view()).toEqual({ phase: "idle", choices: [], pending: false });
  expect(dependencies.ready).not.toHaveBeenCalled(); expect(dependencies.prepare).not.toHaveBeenCalled(); expect(dependencies.run).not.toHaveBeenCalled();
  expect(controller.diagnostics().entries.map(e => e.stage)).toEqual(["runtime.init"]);
});

it.each(["A", "C"] as ProbeKind[])("%sは指定待機後に準備なしで一度だけ実行し成功後も再実行しない", async kind => {
  const { controller, dependencies } = setup();
  expect(controller.schedule(kind, undefined, 3000)).toBe(true);
  expect(controller.view().dueAt).toBe(Date.now() + 3000);
  await vi.advanceTimersByTimeAsync(2999); expect(dependencies.run).not.toHaveBeenCalled();
  await vi.advanceTimersByTimeAsync(1); expect(controller.view()).toMatchObject({phase:"succeeded", pending:false, count:1, kind});
  expect(dependencies.run).toHaveBeenCalledExactlyOnceWith(kind, undefined, expect.any(Function));
  expect(dependencies.prepare).not.toHaveBeenCalled(); expect(controller.schedule(kind, undefined, 0)).toBe(false);
  await controller.prepare(); await vi.advanceTimersByTimeAsync(60000); expect(dependencies.run).toHaveBeenCalledOnce();
  expect(controller.diagnostics().apiCalls).toEqual({preparation:{run:0,sync:0},experiment:{run:1,sync:1}});
});

it.each([-1, 0.5, 30001, NaN, Infinity])("不正な待機時間%sをAPI開始前に拒否", delay => {
  const { controller, dependencies } = setup(); expect(controller.schedule("A", undefined, delay)).toBe(false);
  expect(controller.view().phase).toBe("idle"); expect(dependencies.ready).not.toHaveBeenCalled();
});

it("未知の実験種別を拒否する", () => {
  const { controller } = setup(); expect(controller.schedule("unknown" as ProbeKind, undefined, 0)).toBe(false);
  expect(controller.view().phase).toBe("idle");
});

it("予約キャンセルでもそのページの試験を消費し、Excelを呼ばない", async () => {
  const { controller, dependencies } = setup(); controller.schedule("C", undefined, 1000); controller.cancelSchedule();
  expect(controller.view().phase).toBe("cancelled"); expect(controller.schedule("A", undefined, 0)).toBe(false);
  await controller.prepare(); await vi.advanceTimersByTimeAsync(60000);
  expect(dependencies.ready).not.toHaveBeenCalled(); expect(dependencies.prepare).not.toHaveBeenCalled(); expect(dependencies.run).not.toHaveBeenCalled();
  expect(controller.diagnostics().counts["experiment.cancelledBeforeApi"]).toBe(1);
});

it("Bは一回だけ準備してローカル選択肢のIDを指定し、本試験一回だけ実行", async () => {
  const { controller, dependencies } = setup(); expect(controller.schedule("B", "private-id", 0)).toBe(false);
  await controller.prepare(); await controller.prepare();
  expect(dependencies.prepare).toHaveBeenCalledOnce(); expect(controller.view()).toMatchObject({phase:"ready",kind:"B",choices});
  expect(controller.schedule("B", "foreign-id", 0)).toBe(false); expect(controller.schedule("B", undefined, 0)).toBe(false);
  expect(controller.schedule("A", undefined, 0)).toBe(false); expect(controller.schedule("C", undefined, 0)).toBe(false);
  expect(controller.schedule("B", "private-id", 0)).toBe(true); expect(controller.schedule("B", "private-id", 0)).toBe(false);
  await vi.advanceTimersByTimeAsync(0);
  expect(dependencies.run).toHaveBeenCalledExactlyOnceWith("B", "private-id", expect.any(Function));
  expect(controller.view().phase).toBe("succeeded"); expect(controller.schedule("B", "private-id", 0)).toBe(false);
  expect(controller.diagnostics().apiCalls).toEqual({preparation:{run:1,sync:1},experiment:{run:1,sync:1}});
});

it("準備に失敗した環境でA/Cへ変更したり準備を再試行したりしない", async () => {
  const { controller, dependencies } = setup(); dependencies.prepare.mockImplementation(trace => traced(trace, async () => { throw Object.assign(new Error("private details"),{code:"InvalidOperation"}); }));
  await controller.prepare(); expect(controller.view()).toMatchObject({phase:"failed",errorCode:"InvalidOperation",pending:false});
  expect(controller.schedule("A", undefined, 0)).toBe(false); expect(controller.schedule("C", undefined, 0)).toBe(false);
  await controller.prepare(); expect(dependencies.prepare).toHaveBeenCalledOnce(); expect(dependencies.run).not.toHaveBeenCalled();
});

it("準備が15秒未応答なら試験開始も再取得も禁止し手動マーカーだけ記録できる", async () => {
  const { controller, dependencies } = setup(); const pending = deferred<SheetChoice[]>();
  dependencies.prepare.mockImplementation(trace => traced(trace, () => pending.promise)); const preparation = controller.prepare();
  await vi.advanceTimersByTimeAsync(15000);
  expect(controller.view()).toMatchObject({phase:"stalled",pending:true,waitingAt:"preparation.api.sync"});
  controller.mark("editStart"); controller.mark("escape");
  expect(controller.schedule("A", undefined, 0)).toBe(false); expect(controller.schedule("B", "private-id", 0)).toBe(false);
  await controller.prepare(); expect(dependencies.prepare).toHaveBeenCalledOnce(); expect(dependencies.run).not.toHaveBeenCalled();
  pending.resolve(choices); await preparation; expect(controller.view()).toMatchObject({phase:"settledLate",pending:false,choices:[]});
  expect(controller.schedule("B", "private-id", 0)).toBe(false);
  expect(controller.diagnostics().entries.filter(e => e.source === "manual").map(e => e.stage)).toEqual(["manual.editStart","manual.escape"]);
});

it.each(["resolve", "reject"] as const)("本試験の未応答を警告するだけでAPIを重ねず遅い%s後も再試行しない", async resolution => {
  const { controller, dependencies } = setup(); const pending = deferred<ProbeResult>();
  dependencies.run.mockImplementation((_kind, _id, trace) => traced(trace, () => pending.promise));
  controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(15000);
  expect(controller.view()).toMatchObject({phase:"stalled",pending:true,waitingAt:"experiment.api.sync"});
  controller.mark("editStart"); controller.mark("escape"); controller.cancelSchedule(); await controller.prepare();
  expect(controller.schedule("A", undefined, 0)).toBe(false); expect(controller.schedule("C", undefined, 0)).toBe(false);
  await vi.advanceTimersByTimeAsync(160000); expect(dependencies.run).toHaveBeenCalledOnce(); expect(dependencies.prepare).not.toHaveBeenCalled();
  if (resolution === "resolve") pending.resolve({count:1}); else pending.reject(Object.assign(new Error("private host error"),{code:"InvalidOperation"}));
  await vi.advanceTimersByTimeAsync(0);
  expect(controller.view()).toMatchObject({phase:"settledLate",pending:false}); expect(controller.view().count).toBeUndefined();
  expect(controller.diagnostics().counts["experiment.lateSettled"]).toBe(1); expect(controller.diagnostics().counts["experiment.success"]).toBeUndefined();
  expect(controller.schedule("A", undefined, 0)).toBe(false); await vi.advanceTimersByTimeAsync(60000); expect(dependencies.run).toHaveBeenCalledOnce();
});

it("通常のAPIエラーでも試験は終了し自動・手動で再実行しない", async () => {
  const { controller, dependencies } = setup(); dependencies.run.mockImplementation((_kind,_id,trace) => traced(trace, async () => { throw Object.assign(new Error("private"),{code:"InvalidOperation"}); }));
  controller.schedule("A", undefined, 0); await vi.advanceTimersByTimeAsync(0);
  expect(controller.view()).toMatchObject({phase:"failed",pending:false,errorCode:"InvalidOperation"});
  expect(controller.schedule("C", undefined, 0)).toBe(false); await controller.prepare(); await vi.advanceTimersByTimeAsync(60000);
  expect(dependencies.run).toHaveBeenCalledOnce(); expect(dependencies.prepare).not.toHaveBeenCalled();
});

it("Office準備待ちが期限超過したら遅れて準備完了してもExcel APIを送信しない", async () => {
  const { controller, dependencies } = setup(); const pending = deferred<void>(); dependencies.ready.mockImplementation(() => pending.promise);
  controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(15000);
  expect(controller.view()).toMatchObject({phase:"stalled",pending:true,waitingAt:"experiment.ready"});
  pending.resolve(); await vi.advanceTimersByTimeAsync(0);
  expect(controller.view()).toMatchObject({phase:"settledLate",pending:false}); expect(dependencies.run).not.toHaveBeenCalled();
  expect(controller.diagnostics().counts["experiment.cancelledBeforeApi"]).toBe(1);
});

it("破棄すると予約中のタイマーを止めてAPIを開始しない", async () => {
  const { controller, dependencies } = setup(); controller.schedule("A", undefined, 2000); controller.dispose(); controller.dispose();
  await vi.advanceTimersByTimeAsync(30000); expect(dependencies.ready).not.toHaveBeenCalled(); expect(dependencies.run).not.toHaveBeenCalled();
  expect(controller.view().phase).toBe("disposed"); expect(controller.diagnostics().counts["runtime.dispose"]).toBe(1);
  expect(controller.schedule("A", undefined, 0)).toBe(false);
});

it("Office準備待ちで破棄すると準備完了後もAPIを送らず通知しない", async () => {
  const { controller, dependencies } = setup(); const pending = deferred<void>(); dependencies.ready.mockImplementation(() => pending.promise);
  const listener = vi.fn(); controller.subscribe(listener); controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(0);
  controller.dispose(); const notifications = listener.mock.calls.length; pending.resolve(); await vi.advanceTimersByTimeAsync(30000);
  expect(dependencies.run).not.toHaveBeenCalled(); expect(controller.view()).toMatchObject({phase:"disposed",pending:false});
  expect(listener).toHaveBeenCalledTimes(notifications); expect(controller.diagnostics().counts["experiment.stalled"]).toBeUndefined();
});

it("API未応答中に破棄しても完了したふりをせず結果を画面へ適用しない", async () => {
  const { controller, dependencies } = setup(); const pending = deferred<ProbeResult>(); dependencies.run.mockImplementation((_kind,_id,trace) => traced(trace, () => pending.promise));
  controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(0); controller.dispose();
  expect(controller.view()).toMatchObject({phase:"disposed",pending:true}); pending.resolve({count:9}); await vi.advanceTimersByTimeAsync(0);
  expect(controller.view()).toMatchObject({phase:"disposed",pending:false}); expect(controller.view().count).toBeUndefined();
  expect(dependencies.run).toHaveBeenCalledOnce(); expect(controller.diagnostics().counts["experiment.success"]).toBeUndefined();
});

it("別実行環境を表す別controllerはセッション・待機・ログを共有しない", async () => {
  const first = setup(), second = setup(); const pending = deferred<ProbeResult>(); first.dependencies.run.mockImplementation((_kind,_id,trace) => traced(trace, () => pending.promise));
  first.controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(15000);
  expect(first.controller.view().phase).toBe("stalled"); expect(second.controller.session).not.toBe(first.controller.session);
  expect(second.controller.schedule("A", undefined, 0)).toBe(true); await vi.advanceTimersByTimeAsync(0);
  expect(second.controller.view().phase).toBe("succeeded"); expect(first.controller.view().pending).toBe(true);
  first.controller.mark("escape"); expect(second.controller.diagnostics().counts["manual.escape"]).toBeUndefined();
  pending.resolve({count:1}); await vi.advanceTimersByTimeAsync(0);
});

it("診断は時刻・回数・安全なエラーだけでID・名前・URL・エラー本文を含めない", async () => {
  const { controller, dependencies } = setup(); await controller.prepare();
  dependencies.run.mockImplementation((_kind,_id,trace) => traced(trace, async () => { throw Object.assign(new Error("private-error https://private.example/book"),{code:"InvalidOperation",debugInfo:{errorLocation:"Worksheet.id",fullStatements:["private-id"]}}); }));
  controller.mark("editStart"); await vi.advanceTimersByTimeAsync(20); controller.mark("escape");
  controller.schedule("B", "private-id", 10); await vi.advanceTimersByTimeAsync(10);
  const log = controller.diagnostics(), json = JSON.stringify(log);
  expect(json).not.toMatch(/private|https:/); expect(log.apiCalls).toEqual({preparation:{run:1,sync:1},experiment:{run:1,sync:1}});
  expect(log.entries.find(e => e.stage === "experiment.error")).toMatchObject({code:"InvalidOperation",location:"Worksheet.id",source:"system"});
  const marks = log.entries.filter(e => e.source === "manual"); expect(marks).toHaveLength(2); expect(marks[1].time - marks[0].time).toBe(20);
  expect(marks[0].isoTime).toBe(new Date(marks[0].time).toISOString()); expect(marks[0].elapsedMs).toBeGreaterThanOrEqual(0);
  expect(log.entries.find(e => e.stage === "experiment.api.sync.end")?.durationMs).toBe(7);
});

it("不正なcode/locationや手動マーカー値をログへ入れない", async () => {
  const { controller, dependencies } = setup(); dependencies.run.mockRejectedValue(Object.assign(new Error("private-text"),{code:"private text",debugInfo:{errorLocation:"https://private.example"}}));
  controller.mark("private-marker" as "escape"); controller.schedule("C", undefined, 0); await vi.advanceTimersByTimeAsync(0);
  expect(controller.view().errorCode).toBe("UnknownError"); expect(JSON.stringify(controller.diagnostics())).not.toContain("private");
});

it("外部がview/diagnosticsを変更しても内部の選択肢・履歴を変更できない", async () => {
  const { controller } = setup(); await controller.prepare(); const view = controller.view(); view.choices[0].id = "foreign"; view.phase = "failed";
  const log = controller.diagnostics(); log.entries[0].stage = "foreign"; log.counts["runtime.init"] = 999;
  expect(controller.view().phase).toBe("ready"); expect(controller.schedule("B", "foreign", 0)).toBe(false);
  expect(controller.diagnostics().entries[0].stage).toBe("runtime.init"); expect(controller.diagnostics().counts["runtime.init"]).toBe(1);
});
