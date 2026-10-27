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
afterEach(cleanup);
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
it("検索中の無効理由を表示しクリアで解除",async()=>{const hook=await ready();act(()=>hook.result.current.setQuery("Sheet"));expect(hook.result.current.moveDisabledReason).toContain("検索中");act(()=>hook.result.current.setQuery(""));expect(hook.result.current.moveDisabledReason).toBeNull();});
it("移動失敗でも再取得した状態を表示しエラーを保持",async()=>{const hook=await ready();const actual=workbook();actual.sheets[0].name="current";vi.mocked(reorderWorksheet).mockResolvedValue({state:actual,error:"移動失敗"});await act(()=>hook.result.current.move(workbook(),"b",0,0));expect(hook.result.current.sheets[0].name).toBe("current");expect(hook.result.current.error).toBe("移動失敗");});
it("移動の連打を抑止し処理中busy",async()=>{const hook=await ready();let resolve!:(value:{state:WorkbookState})=>void;vi.mocked(reorderWorksheet).mockImplementation(()=>new Promise(r=>{resolve=r;}));let first!:Promise<void>;act(()=>{first=hook.result.current.move(workbook(),"b",0,0);void hook.result.current.move(workbook(),"b",0,0);});expect(hook.result.current.busy).toBe(true);expect(reorderWorksheet).toHaveBeenCalledOnce();await act(async()=>{resolve({state:workbook()});await first;});expect(hook.result.current.busy).toBe(false);});
it("登録が完了する前に閉じても完了後に解除",async()=>{let finish!:(release:()=>Promise<void>)=>void;vi.mocked(worksheetEvents.subscribe).mockImplementation(()=>new Promise(r=>{finish=r;}));const hook=renderHook(()=>useWorksheets());await waitFor(()=>expect(worksheetEvents.subscribe).toHaveBeenCalledOnce());hook.unmount();await act(async()=>finish(dispose));expect(dispose).toHaveBeenCalledOnce();});

it("自動更新エラーは次イベントの取得成功で解消",async()=>{const hook=await ready();vi.mocked(readWorkbookState).mockRejectedValueOnce(new Error("temporary"));event();await waitFor(()=>expect(hook.result.current.error).toContain("temporary"));event();await waitFor(()=>expect(hook.result.current.error).toBeNull());expect(hook.result.current.sheets).toHaveLength(3);});
it("フック経由でも検索中は移動を拒否",async()=>{const hook=await ready();act(()=>hook.result.current.setQuery("Sheet2"));await act(()=>hook.result.current.move(workbook(),"b",0,0));expect(reorderWorksheet).not.toHaveBeenCalled();expect(hook.result.current.error).toContain("検索");});
