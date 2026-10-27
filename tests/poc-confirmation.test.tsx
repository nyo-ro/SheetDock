// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { Panel } from "../src/poc/Panel";
import { ensureExcelReady } from "../src/excel/environment";
import { applyAction, observeBook } from "../src/poc/officeAdapter";
import { book, memoryAdapter } from "./poc-fixtures";
vi.mock("../src/excel/environment",()=>({ensureExcelReady:vi.fn(async()=>{})}));
vi.mock("../src/poc/officeAdapter",()=>({observeBook:vi.fn(),applyAction:vi.fn()}));
vi.mock("../src/poc/events",()=>({eventSets:{},listenEvents:()=>()=>{},eventLifecycle:{setEnabled:vi.fn(async()=>{})}}));
let adapter: ReturnType<typeof memoryAdapter>;
beforeEach(()=>{
  vi.clearAllMocks(); vi.mocked(ensureExcelReady).mockResolvedValue(undefined);
  const initial=book(); initial.name="Book 1.xlsx"; initial.sheets=initial.sheets.slice(0,3); initial.totalSheets=3;
  initial.sheets.forEach((sheet,index)=>{sheet.name=`Sheet${index+1}`;});
  adapter=memoryAdapter(initial);
  vi.mocked(observeBook).mockImplementation(adapter.read);
  vi.mocked(applyAction).mockImplementation(adapter.apply);
  Object.defineProperty(window,"matchMedia",{configurable:true,writable:true,value:vi.fn(query=>({matches:false,media:query,onchange:null,addListener:vi.fn(),removeListener:vi.fn(),addEventListener:vi.fn(),removeEventListener:vi.fn(),dispatchEvent:vi.fn()}))});
  // Officeで非対応のネイティブ確認に戻ったらテストを失敗させる。
  vi.spyOn(window,"confirm").mockImplementation(()=>{throw new Error("Function window.confirm is not supported.");});
  vi.spyOn(window,"alert").mockImplementation(()=>{throw new Error("native alert forbidden");});
  vi.spyOn(window,"prompt").mockImplementation(()=>{throw new Error("native prompt forbidden");});
});
afterEach(()=>{cleanup();vi.restoreAllMocks();});
async function prepared() {
  const user=userEvent.setup();render(<Panel />);
  await user.click(screen.getByRole("button",{name:"観測・再取得（変更なし）"}));
  await user.click(screen.getByLabelText("使い捨て検証ブックで変更操作を行う"));
  await user.click(screen.getByLabelText(/通常ワークシートのみのブックと自分で確認/));
  await user.type(screen.getByLabelText(/確認用ブック名/),"Book 1.xlsx");
  await user.type(screen.getByLabelText(/対象ID/),"b");
  await user.selectOptions(screen.getByLabelText("操作"),"move");
  await user.click(screen.getByRole("button",{name:"計画を作成（変更なし）"}));
  return user;
}
async function opened() {
  const user=await prepared();await user.click(screen.getByRole("button",{name:"確認して実行"}));
  return {user,dialog:await screen.findByRole("dialog",{name:"検証操作の実行確認"})};
}
it("確認にブック・対象・操作・期待順序を表示し、まだ変更しない",async()=>{
  const {dialog}=await opened();expect(dialog.textContent).toContain("Book 1.xlsx");
  expect(dialog.textContent).toContain("Sheet2（ID: b） を位置 1 へ移動");
  expect(dialog.textContent).toContain("Sheet2 → Sheet1 → Sheet3");
  expect(applyAction).not.toHaveBeenCalled();expect(window.confirm).not.toHaveBeenCalled();
});
it("キャンセルすると変更APIを呼ばず計画を残す",async()=>{
  const {user,dialog}=await opened();await user.click(within(dialog).getByRole("button",{name:"キャンセル"}));
  expect(screen.queryByRole("dialog")).toBeNull();expect(applyAction).not.toHaveBeenCalled();
  expect((screen.getByRole("button",{name:"確認して実行"}) as HTMLButtonElement).disabled).toBe(false);
});
it("Escapeで閉じても実行しない",async()=>{
  const {user}=await opened();await user.keyboard("{Escape}");
  await waitFor(()=>expect(screen.queryByRole("dialog")).toBeNull());expect(applyAction).not.toHaveBeenCalled();
});
it("確認を閉じてペインを破棄しても変更しない",async()=>{
  await opened();cleanup();expect(applyAction).not.toHaveBeenCalled();
});
it("明示実行でのみ再取得・変更・結果照合へ進む",async()=>{
  const {user,dialog}=await opened();expect(observeBook).toHaveBeenCalledTimes(1);
  await user.click(within(dialog).getByRole("button",{name:"実行"}));
  await screen.findByText(/"orderMatches": true/);
  expect(observeBook).toHaveBeenCalledTimes(3);expect(applyAction).toHaveBeenCalledTimes(1);
  expect(vi.mocked(observeBook).mock.invocationCallOrder[1]).toBeLessThan(vi.mocked(applyAction).mock.invocationCallOrder[0]);
  expect((await adapter.read()).sheets.map(sheet=>sheet.name)).toEqual(["Sheet2","Sheet1","Sheet3"]);
});
it.each(["name","position","visibility","readOnly","structureProtected","sheetProtection"])("確認待ち中の%s変更を再検証で拒否",async property=>{
  const {user,dialog}=await opened();
  adapter.change(state=>{
    if(property==="name")state.sheets[1].name="Changed";
    if(property==="position"){[state.sheets[0],state.sheets[1]]=[state.sheets[1],state.sheets[0]];state.sheets.forEach((sheet,index)=>{sheet.position=index;});}
    if(property==="visibility")state.sheets[1].visibility="Hidden";
    if(property==="readOnly")state.readOnly=true;
    if(property==="structureProtected")state.structureProtected=true;
    if(property==="sheetProtection")state.sheets[1].protected=true;
  });
  await user.click(within(dialog).getByRole("button",{name:"実行"}));
  const result=await screen.findByText(/"status": "failed"/);
  expect(result.textContent).toContain("変化しました");expect(applyAction).not.toHaveBeenCalled();
});
it("同じ実行ボタンの連打でも一度だけ処理し、実行中は操作を無効化",async()=>{
  const {dialog}=await opened();let release!:()=>void;
  const perform=adapter.apply.getMockImplementation()!;
  vi.mocked(applyAction).mockImplementation(async action=>{await new Promise<void>(resolve=>{release=resolve;});await perform(action);});
  const execute=within(dialog).getByRole("button",{name:"実行"});
  act(()=>{fireEvent.click(execute);fireEvent.click(execute);});
  await waitFor(()=>expect(applyAction).toHaveBeenCalledTimes(1));
  expect((screen.getByRole("button",{name:"観測・再取得（変更なし）"}) as HTMLButtonElement).disabled).toBe(true);
  expect((screen.getByRole("button",{name:"確認して実行"}) as HTMLButtonElement).disabled).toBe(true);
  await act(async()=>{release();});await screen.findByText(/"orderMatches": true/);
  expect(applyAction).toHaveBeenCalledTimes(1);
});
it("初期化エラーをPoCのエラー領域へ表示し書き込まない",async()=>{
  const {user,dialog}=await opened();vi.mocked(ensureExcelReady).mockRejectedValueOnce(new Error("接続を確認できません"));
  await user.click(within(dialog).getByRole("button",{name:"実行"}));
  expect((await screen.findByRole("alert")).textContent).toContain("接続を確認できません");expect(applyAction).not.toHaveBeenCalled();
});
it("送信後のAPIエラーは既存結果へ状態不明として表示",async()=>{
  const {user,dialog}=await opened();vi.mocked(applyAction).mockRejectedValueOnce(new Error("AccessDenied mock"));
  await user.click(within(dialog).getByRole("button",{name:"実行"}));
  expect((await screen.findByText(/"status": "unknown"/)).textContent).toContain("AccessDenied mock");
});
it("ブック名完全一致と2つの確認を省略できない",async()=>{
  const user=await prepared();
  await user.clear(screen.getByLabelText(/確認用ブック名/));await user.type(screen.getByLabelText(/確認用ブック名/),"Book 1");
  expect((screen.getByRole("button",{name:"確認して実行"}) as HTMLButtonElement).disabled).toBe(true);
  await user.type(screen.getByLabelText(/確認用ブック名/),".xlsx");
  await user.click(screen.getByLabelText(/通常ワークシートのみのブックと自分で確認/));
  await user.click(screen.getByRole("button",{name:"計画を作成（変更なし）"}));
  expect((screen.getByRole("button",{name:"確認して実行"}) as HTMLButtonElement).disabled).toBe(true);
  await user.click(screen.getByLabelText("使い捨て検証ブックで変更操作を行う"));
  expect((screen.getByRole("button",{name:"確認して実行"}) as HTMLButtonElement).disabled).toBe(true);expect(applyAction).not.toHaveBeenCalled();
});

it("名前選択だけで複数移動の計画・確認・実行ができる",async()=>{
 const user=await prepared();await user.clear(screen.getByLabelText(/対象ID/));
 await user.click(screen.getByLabelText("対象シート Sheet1"));await user.click(screen.getByLabelText("対象シート Sheet3"));
 await user.selectOptions(screen.getByLabelText("移動先を名前で選択"),"3");
 await user.click(screen.getByRole("button",{name:"計画を作成（変更なし）"}));await user.click(screen.getByRole("button",{name:"確認して実行"}));
 const dialog=await screen.findByRole("dialog");expect(dialog.textContent).toContain("Sheet2 → Sheet1 → Sheet3");expect(applyAction).not.toHaveBeenCalled();
 await user.click(within(dialog).getByRole("button",{name:"実行"}));await screen.findByText(/"orderMatches": true/);
 expect((await adapter.read()).sheets.map(s=>s.name)).toEqual(["Sheet2","Sheet1","Sheet3"]);
});
