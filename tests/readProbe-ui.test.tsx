// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react";
import { ReadProbePanel } from "../src/dev/readProbe/Panel";
import { prepareKnownSheets, runReadProbe, ProbeTrace } from "../src/dev/readProbe/api";
import { ensureExcelReady } from "../src/excel/environment";
import { reloadInMode } from "../src/dev/readProbe/mode";
vi.mock("../src/dev/readProbe/api", () => ({ prepareKnownSheets: vi.fn(), runReadProbe: vi.fn() }));
vi.mock("../src/excel/environment", () => ({ ensureExcelReady: vi.fn() }));
vi.mock("../src/dev/readProbe/mode", () => ({ reloadInMode: vi.fn() }));
const button = (name: string) => screen.getByRole("button", { name });
const disabled = (element: HTMLElement) => element.matches(":disabled");
async function elapsed(ms: number) { await act(async () => { await vi.advanceTimersByTimeAsync(ms); }); }
async function traced<T>(trace: ProbeTrace, work: () => Promise<T>): Promise<T> {
  trace("api.run.start"); trace("api.sync.start");
  try { return await work(); } catch (e) { trace("api.sync.error", e); throw e; }
  finally { trace("api.sync.end", undefined, 4); trace("api.run.end", undefined, 5); }
}
beforeEach(() => {
  vi.resetAllMocks(); vi.useFakeTimers();
  vi.mocked(ensureExcelReady).mockResolvedValue(undefined);
  vi.mocked(prepareKnownSheets).mockImplementation(trace => traced(trace, async () => [{id:"private-a",name:"対象A",position:0},{id:"private-b",name:"対象B",position:1}]));
  vi.mocked(runReadProbe).mockImplementation((_kind,_id,trace) => traced(trace, async () => ({count:1})));
});
afterEach(() => { cleanup(); vi.useRealTimers(); });

it("画面を開いて待機してもExcel取得やイベント同期を起動しない", async () => {
  render(<ReadProbePanel />); await elapsed(60000);
  expect(screen.getByText(/通常UI・PoC・イベント購読は起動していません/)).toBeTruthy();
  expect(ensureExcelReady).not.toHaveBeenCalled(); expect(prepareKnownSheets).not.toHaveBeenCalled(); expect(runReadProbe).not.toHaveBeenCalled();
});
it("Cを明示予約した時だけ実行し連打・試験変更・成功後の再実行を禁止", async () => {
  render(<ReadProbePanel />); fireEvent.click(screen.getByRole("radio", {name:"C：アクティブシートIDのみ"}));
  fireEvent.click(button("読み取りを予約実行")); fireEvent.click(button("読み取りを予約実行"));
  expect(disabled(screen.getByRole("radio", {name:"A：ワークシート一覧のみ"}))).toBe(true);
  await elapsed(4999); expect(runReadProbe).not.toHaveBeenCalled(); await elapsed(1);
  expect(runReadProbe).toHaveBeenCalledExactlyOnceWith("C", undefined, expect.any(Function)); expect(prepareKnownSheets).not.toHaveBeenCalled();
  expect(screen.getByText(/状態: 取得成功/)).toBeTruthy(); expect(disabled(button("読み取りを予約実行"))).toBe(true);
});
it("予約取消は変更・取得APIを呼ばず、新しい環境への移動だけを許可", async () => {
  render(<ReadProbePanel />); fireEvent.click(button("読み取りを予約実行")); fireEvent.click(button("予約を取り消す")); await elapsed(30000);
  expect(runReadProbe).not.toHaveBeenCalled(); expect(ensureExcelReady).not.toHaveBeenCalled(); expect(disabled(button("読み取りを予約実行"))).toBe(true);
  fireEvent.click(button("新しい実行環境へ（再読み込み）")); expect(reloadInMode).toHaveBeenCalledWith(true);
});
it("Bは準備完了後のシート名選択から内部IDを渡し、準備と本試験の回数を分ける", async () => {
  render(<ReadProbePanel />); fireEvent.click(screen.getByRole("radio",{name:"B：既知IDのワークシート"}));
  expect(disabled(button("読み取りを予約実行"))).toBe(true);
  fireEvent.click(button("Bの対象を準備")); await elapsed(0); expect(prepareKnownSheets).toHaveBeenCalledOnce();
  expect(screen.getByRole("option",{name:"2：対象B"})).toBeTruthy();
  fireEvent.change(screen.getByLabelText("対象シート"), {target:{value:"private-b"}});
  fireEvent.change(screen.getByLabelText("開始までの秒数"), {target:{value:"0"}}); fireEvent.click(button("読み取りを予約実行")); await elapsed(0);
  expect(runReadProbe).toHaveBeenCalledExactlyOnceWith("B", "private-b", expect.any(Function));
  fireEvent.click(button("診断ログを表示・更新")); const json=(screen.getByLabelText("診断ログJSON") as HTMLTextAreaElement).value;
  expect(JSON.parse(json).apiCalls).toEqual({preparation:{run:1,sync:1},experiment:{run:1,sync:1}});
  expect(json).not.toContain("private"); expect(json).not.toContain("対象B");
});
it("API無応答中は次試験・画面復帰を無効化し、手動時刻と診断保存は動作する", async () => {
  let finish!: (value:{count:number}) => void;
  vi.mocked(runReadProbe).mockImplementation((_kind,_id,trace) => traced(trace, () => new Promise(resolve => {finish=resolve;})));
  render(<ReadProbePanel />); fireEvent.change(screen.getByLabelText("開始までの秒数"), {target:{value:"0"}}); fireEvent.click(button("読み取りを予約実行")); await elapsed(15000);
  expect(screen.getByRole("alert").textContent).toContain("15秒以上");
  for (const name of ["読み取りを予約実行","新しい実行環境へ（再読み込み）","通常画面へ（再読み込み）"]) { expect(disabled(button(name))).toBe(true); fireEvent.click(button(name)); }
  fireEvent.click(button("名前編集開始を記録（手動）")); fireEvent.click(button("Esc取消を記録（手動）")); fireEvent.click(button("診断ログを表示・更新"));
  const log=JSON.parse((screen.getByLabelText("診断ログJSON") as HTMLTextAreaElement).value);
  expect(log.pending).toBe(true); expect(log.waitingAt).toBe("experiment.api.sync"); expect(log.counts["manual.editStart"]).toBe(1); expect(log.counts["manual.escape"]).toBe(1);
  expect(log.counts["experiment.api.sync.end"]).toBeUndefined(); expect(runReadProbe).toHaveBeenCalledOnce(); expect(reloadInMode).not.toHaveBeenCalled();
  await act(async()=>finish({count:1})); expect(screen.getByText(/状態: 期限超過後に応答/)).toBeTruthy(); expect(disabled(button("読み取りを予約実行"))).toBe(true);
});
it("エラーの生本文を表示・診断に出さず、安全なコードだけを表示する", async () => {
  vi.mocked(runReadProbe).mockImplementation((_kind,_id,trace) => traced(trace,async()=>{throw Object.assign(new Error("secret sheet/book content"),{code:"InvalidOperation",debugInfo:{errorLocation:"Worksheet.id"}});}));
  render(<ReadProbePanel />); fireEvent.change(screen.getByLabelText("開始までの秒数"),{target:{value:"0"}}); fireEvent.click(button("読み取りを予約実行")); await elapsed(0);
  expect(screen.getByRole("alert").textContent).toContain("InvalidOperation"); expect(document.body.textContent).not.toContain("secret");
  fireEvent.click(button("診断ログを表示・更新")); const json=(screen.getByLabelText("診断ログJSON") as HTMLTextAreaElement).value;
  expect(json).not.toContain("secret"); expect(json).toContain("Worksheet.id"); expect(disabled(button("読み取りを予約実行"))).toBe(true);
});
it("画面終了時は予約を取り消し、後からAPIを実行しない", async () => {
  const mounted=render(<ReadProbePanel />); fireEvent.click(button("読み取りを予約実行")); mounted.unmount(); await elapsed(60000); expect(runReadProbe).not.toHaveBeenCalled();
});
it("通常画面への復帰もモード変更と新しいページ起動を要求する", () => {
  render(<ReadProbePanel />); fireEvent.click(button("通常画面へ（再読み込み）")); expect(reloadInMode).toHaveBeenCalledExactlyOnceWith(false);
});
