// @vitest-environment jsdom
import { Mock, afterEach, beforeEach, expect, it, vi } from "vitest";
import { act, cleanup, renderHook, waitFor } from "@testing-library/react";
import { useWorksheets } from "../src/state/useWorksheets";
import { readWorkbookState } from "../src/excel/workbookState";
import { worksheetEvents, EventRecord } from "../src/excel/worksheetEvents";
import { reorderWorksheet } from "../src/excel/reorderWorksheet";
import { WorkbookState } from "../src/domain/workbookState";
import { workbook } from "./production-fixture";
vi.mock("../src/excel/environment",()=>({ensureExcelReady:async()=>{},errorMessage:(e:Error)=>e.message}));
vi.mock("../src/excel/workbookState",()=>({readWorkbookState:vi.fn()}));
vi.mock("../src/excel/worksheetEvents",()=>({worksheetEvents:{subscribe:vi.fn(),rebind:vi.fn()}}));
vi.mock("../src/excel/reorderWorksheet",()=>({reorderWorksheet:vi.fn()}));
vi.mock("../src/excel/worksheets",()=>({activateWorksheet:vi.fn(async()=>{})}));
let emit:(e:EventRecord)=>void;let dispose:Mock<()=>Promise<void>>;
beforeEach(()=>{vi.resetAllMocks();vi.mocked(readWorkbookState).mockResolvedValue(workbook());dispose=vi.fn(async()=>{});vi.mocked(worksheetEvents.subscribe).mockImplementation(async callback=>{emit=callback;return dispose;});vi.mocked(worksheetEvents.rebind).mockResolvedValue(undefined);});
afterEach(()=>{cleanup();vi.useRealTimers();});
async function ready(){const hook=renderHook(()=>useWorksheets());await waitFor(()=>expect(hook.result.current.busy).toBe(false));return hook;}
const event=(type="onNameChanged")=>act(()=>emit({time:"test",type,args:{}}));
it("通常UI起動で登録・取得し終了で解除",async()=>{const hook=await ready();expect(worksheetEvents.subscribe).toHaveBeenCalledTimes(1);expect(hook.result.current.sheets).toHaveLength(3);hook.unmount();await waitFor(()=>expect(dispose).toHaveBeenCalledOnce());});
it("大量イベントを集約して一覧とアクティブを更新",async()=>{
 const hook=await ready();const updated=workbook();updated.sheets[1].name="Renamed";updated.activeId="b";updated.sheets.forEach(s=>{s.isActive=s.id==="b";});vi.mocked(readWorkbookState).mockResolvedValue(updated);
 act(()=>{for(let i=0;i<100;i++)emit({time:"test",type:"onNameChanged",args:{}});});
 await waitFor(()=>expect(hook.result.current.sheets[1].name).toBe("Renamed"));expect(hook.result.current.sheets[1].isActive).toBe(true);expect(readWorkbookState).toHaveBeenCalledTimes(2);
});
it("更新中のイベントを後続1回へ持ち越す",async()=>{
 await ready();let resolve!:(state:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{resolve=r;}));event();await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledTimes(2));event();event();await act(async()=>resolve(workbook()));await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledTimes(3));
});
it("ブックのアクティブ化で再登録し古いドラッグ世代を無効化",async()=>{
 const hook=await ready();const old=hook.result.current.dragEpoch();event("workbookActivated");expect(hook.result.current.dragEpoch()).not.toBe(old);expect(hook.result.current.snapshot).toBeUndefined();await waitFor(()=>expect(worksheetEvents.rebind).toHaveBeenCalledOnce());await waitFor(()=>expect(hook.result.current.snapshot).toBeDefined());
});
it("取得途中のブック切替で旧結果を採用しない",async()=>{
 const hook=await ready();let resolve!:(state:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{resolve=r;}));event();await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledTimes(2));event("workbookActivated");
 const stale=workbook();stale.sheets[0].name="stale";await act(async()=>resolve(stale));expect(hook.result.current.snapshot).toBeUndefined();await waitFor(()=>expect(hook.result.current.sheets[0]?.name).toBe("Sheet1"));
});
it("ブック識別情報変化も再登録する",async()=>{const hook=await ready();const next=workbook();next.bookKey="book-2";vi.mocked(readWorkbookState).mockResolvedValue(next);event();await waitFor(()=>expect(hook.result.current.snapshot?.bookKey).toBe("book-2"));expect(worksheetEvents.rebind).toHaveBeenCalledOnce();});
it("登録失敗後も一覧を表示し手動更新で再接続",async()=>{vi.mocked(worksheetEvents.subscribe).mockRejectedValueOnce(new Error("offline"));const hook=await ready();expect(hook.result.current.syncError).toContain("offline");expect(hook.result.current.sheets).toHaveLength(3);await act(()=>hook.result.current.refresh());expect(hook.result.current.syncError).toBeNull();expect(worksheetEvents.subscribe).toHaveBeenCalledTimes(2);});
it("自動更新失敗は古い一覧を消し、手動更新で回復",async()=>{const hook=await ready();vi.mocked(readWorkbookState).mockRejectedValueOnce(new Error("offline"));event();await waitFor(()=>expect(hook.result.current.error).toContain("offline"));expect(hook.result.current.sheets).toHaveLength(0);await act(()=>hook.result.current.refresh());expect(hook.result.current.sheets).toHaveLength(3);expect(hook.result.current.error).toBeNull();});
it("検索中も移動を許可しクリア後も安全判定を維持",async()=>{const hook=await ready();act(()=>hook.result.current.setQuery("Sheet"));expect(hook.result.current.moveDisabledReason).toBeNull();act(()=>hook.result.current.setQuery(""));expect(hook.result.current.moveDisabledReason).toBeNull();});
it("移動失敗でも再取得した状態を表示しエラーを保持",async()=>{const hook=await ready();const actual=workbook();actual.sheets[0].name="current";vi.mocked(reorderWorksheet).mockResolvedValue({state:actual,error:"移動失敗"});await act(()=>hook.result.current.move(workbook(),"b",0,0));expect(hook.result.current.sheets[0].name).toBe("current");expect(hook.result.current.error).toBe("移動失敗");});
it("移動の連打を抑止し処理中busy",async()=>{const hook=await ready();let resolve!:(value:{state:WorkbookState})=>void;vi.mocked(reorderWorksheet).mockImplementation(()=>new Promise(r=>{resolve=r;}));let first!:Promise<void>;act(()=>{first=hook.result.current.move(workbook(),"b",0,0);void hook.result.current.move(workbook(),"b",0,0);});expect(hook.result.current.busy).toBe(true);expect(reorderWorksheet).toHaveBeenCalledOnce();await act(async()=>{resolve({state:workbook()});await first;});expect(hook.result.current.busy).toBe(false);});
it("登録が完了する前に閉じても完了後に解除",async()=>{let finish!:(release:()=>Promise<void>)=>void;vi.mocked(worksheetEvents.subscribe).mockImplementation(()=>new Promise(r=>{finish=r;}));const hook=renderHook(()=>useWorksheets());await waitFor(()=>expect(worksheetEvents.subscribe).toHaveBeenCalledOnce());hook.unmount();await act(async()=>finish(dispose));expect(dispose).toHaveBeenCalledOnce();});

it("自動更新エラーは次イベントの取得成功で解消",async()=>{const hook=await ready();vi.mocked(readWorkbookState).mockRejectedValueOnce(new Error("temporary"));event();await waitFor(()=>expect(hook.result.current.error).toContain("temporary"));event();await waitFor(()=>expect(hook.result.current.error).toBeNull());expect(hook.result.current.sheets).toHaveLength(3);});
it("検索中の移動後も検索条件を維持",async()=>{const hook=await ready();act(()=>hook.result.current.setQuery("Sheet2"));vi.mocked(reorderWorksheet).mockResolvedValue({state:workbook()});await act(()=>hook.result.current.move(workbook(),"b",0,0));expect(reorderWorksheet).toHaveBeenCalledOnce();expect(hook.result.current.query).toBe("Sheet2");});

it("初回取得中の名前変更を直列化し後続一回に集約",async()=>{
 let finish!:(state:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{finish=r;}));
 const hook=renderHook(()=>useWorksheets());await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledOnce());event();event();
 await act(async()=>{await new Promise(r=>setTimeout(r,220));});expect(readWorkbookState).toHaveBeenCalledTimes(1);
 await act(async()=>finish(workbook()));await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledTimes(2));hook.unmount();
});

it("終了後の旧イベントを無視し再表示で一度だけ購読する",async()=>{
 const first=await ready();const oldEmit=emit;first.unmount();const second=await ready();expect(worksheetEvents.subscribe).toHaveBeenCalledTimes(2);expect(dispose).toHaveBeenCalledOnce();
 act(()=>oldEmit({time:"old",type:"onNameChanged",args:{}}));await act(async()=>{await new Promise(r=>setTimeout(r,200));});expect(readWorkbookState).toHaveBeenCalledTimes(2);second.unmount();
});

it("再登録中の新しいブックイベントを取りこぼさない",async()=>{
 await ready();let finish!:()=>void;vi.mocked(worksheetEvents.rebind).mockImplementationOnce(()=>new Promise<void>(r=>{finish=r;}));event("workbookActivated");await waitFor(()=>expect(worksheetEvents.rebind).toHaveBeenCalledOnce());event("workbookActivated");await act(async()=>finish());await waitFor(()=>expect(worksheetEvents.rebind).toHaveBeenCalledTimes(2));
});

it("取得途中の連続改名の旧結果を捨て最終名へ追従",async()=>{
 const hook=await ready();let finish!:(s:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{finish=r;}));event();await waitFor(()=>expect(readWorkbookState).toHaveBeenCalledTimes(2));
 const newest=workbook();newest.sheets[0].name="final";vi.mocked(readWorkbookState).mockResolvedValue(newest);for(let i=0;i<10;i++)event();
 const stale=workbook();stale.sheets[0].name="stale";await act(async()=>finish(stale));expect(hook.result.current.sheets[0].name).not.toBe("stale");await waitFor(()=>expect(hook.result.current.sheets[0].name).toBe("final"));expect(readWorkbookState).toHaveBeenCalledTimes(3);
});
it("初回API無応答を15秒で明示し、遅い旧結果を適用せず一度回復取得",async()=>{
 vi.useFakeTimers();let finish!:(s:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{finish=r;}));const hook=renderHook(()=>useWorksheets());
 await act(async()=>{await vi.advanceTimersByTimeAsync(15000);});expect(hook.result.current.busy).toBe(false);expect(hook.result.current.stalled).toBe(true);expect(hook.result.current.ready).toBe(false);expect(hook.result.current.error).toContain("15秒");
 await act(()=>hook.result.current.refresh());expect(readWorkbookState).toHaveBeenCalledOnce();
 const stale=workbook();stale.sheets[0].name="expired";await act(async()=>finish(stale));expect(hook.result.current.sheets.some(s=>s.name==="expired")).toBe(false);
 await act(async()=>{await vi.advanceTimersByTimeAsync(150);});expect(hook.result.current.sheets[0].name).toBe("Sheet1");expect(readWorkbookState).toHaveBeenCalledTimes(2);expect(hook.result.current.error).toBeNull();
});
it("応答遅延が続いてもイベントなしの自動回復は一度に制限",async()=>{
 vi.useFakeTimers();const finishes:Array<(s:WorkbookState)=>void>=[];vi.mocked(readWorkbookState).mockImplementation(()=>new Promise(r=>finishes.push(r)));const hook=renderHook(()=>useWorksheets());
 await act(async()=>{await vi.advanceTimersByTimeAsync(15000);});await act(async()=>finishes[0](workbook()));await act(async()=>{await vi.advanceTimersByTimeAsync(15150);});
 await act(async()=>finishes[1](workbook()));await act(async()=>{await vi.advanceTimersByTimeAsync(60000);});expect(readWorkbookState).toHaveBeenCalledTimes(2);expect(hook.result.current.busy).toBe(false);expect(hook.result.current.error).toBeTruthy();expect(hook.result.current.moveDisabledReason).toContain("確認できません");
});
it("初回API拒否でbusyを解除し手動で回復",async()=>{vi.mocked(readWorkbookState).mockRejectedValueOnce(new Error("host failure"));const hook=await ready();expect(hook.result.current.error).toContain("host failure");await act(()=>hook.result.current.refresh());expect(hook.result.current.sheets).toHaveLength(3);expect(hook.result.current.error).toBeNull();});
it("終了後に無応答監視の通知や旧取得結果を新ペインへ適用しない",async()=>{
 vi.useFakeTimers();let finish!:(s:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{finish=r;}));const first=renderHook(()=>useWorksheets());await act(async()=>{await vi.advanceTimersByTimeAsync(0);});first.unmount();
 const second=renderHook(()=>useWorksheets());await act(async()=>{await vi.advanceTimersByTimeAsync(20000);});const stale=workbook();stale.sheets[0].name="old-pane";await act(async()=>finish(stale));expect(second.result.current.stalled).toBe(false);expect(second.result.current.sheets[0].name).toBe("Sheet1");
});

it("手動更新中の無応答でもUIを応答待ちへ移し完了後に復旧",async()=>{
 const hook=await ready();vi.useFakeTimers();let finish!:(s:WorkbookState)=>void;vi.mocked(readWorkbookState).mockImplementationOnce(()=>new Promise(r=>{finish=r;}));let manual!:Promise<void>;act(()=>{manual=hook.result.current.refresh();});
 await act(async()=>{await vi.advanceTimersByTimeAsync(15000);});expect(hook.result.current.busy).toBe(false);expect(hook.result.current.operationBlocked).toBe(true);await act(async()=>{finish(workbook());await manual;});await act(async()=>{await vi.advanceTimersByTimeAsync(150);});expect(hook.result.current.error).toBeNull();expect(hook.result.current.sheets).toHaveLength(3);
});

it("一覧のみの退避結果は警告と移動不可を表示し次回成功で回復",async()=>{
 const fallback=workbook();fallback.activeId=null;fallback.sheets.forEach(s=>{s.isActive=false;});fallback.readIssue="一覧のみ再取得";vi.mocked(readWorkbookState).mockResolvedValueOnce(fallback);const hook=await ready();expect(hook.result.current.error).toBe("一覧のみ再取得");expect(hook.result.current.moveDisabledReason).toContain("アクティブ");expect(hook.result.current.sheets).toHaveLength(3);await act(()=>hook.result.current.refresh());expect(hook.result.current.error).toBeNull();expect(hook.result.current.sheets[0].isActive).toBe(true);
});

it("通常表示の総数が未確認でも一覧・検索を保ちドロップ時の厳密な処理へ渡す",async()=>{
 const display=workbook();display.totalSheets=null;display.readPurpose="display";vi.mocked(readWorkbookState).mockResolvedValue(display);
 const hook=await ready();expect(hook.result.current.hasSnapshot).toBe(true);expect(hook.result.current.sheets).toHaveLength(3);expect(hook.result.current.moveDisabledReason).toBeNull();
 act(()=>hook.result.current.setQuery("Sheet2"));vi.mocked(reorderWorksheet).mockResolvedValue({error:"総数確認失敗",state:{...workbook(),totalSheets:null}});
 await act(()=>hook.result.current.move(display,"b",0,0));expect(reorderWorksheet).toHaveBeenCalledExactlyOnceWith(display,"b",0,expect.any(Function));expect(hook.result.current.query).toBe("Sheet2");expect(hook.result.current.error).toBe("総数確認失敗");expect(hook.result.current.moveDisabledReason).toContain("混在");
});
