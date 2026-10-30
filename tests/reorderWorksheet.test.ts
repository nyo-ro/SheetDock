import { Mock, beforeEach, afterEach, expect, it, vi } from "vitest";
import { reorderWorksheet } from "../src/excel/reorderWorksheet";
import { loadWorkbookState } from "../src/excel/workbookState";
import { moveRestriction, dragRestriction } from "../src/domain/workbookState";
import { workbook } from "./production-fixture";
vi.mock("../src/excel/workbookState",()=>({loadWorkbookState:vi.fn()}));
let current=workbook();let write:Mock<(id:string,position:number)=>void>;let sync:Mock<()=>Promise<void>>;
beforeEach(()=>{
 current=workbook();vi.mocked(loadWorkbookState).mockReset().mockImplementation(async()=>structuredClone(current));
 write=vi.fn((id:string,position:number)=>{const from=current.sheets.findIndex(s=>s.id===id);const [sheet]=current.sheets.splice(from,1);current.sheets.splice(position,0,sheet);current.sheets.forEach((s,i)=>{s.position=i;});});
 sync=vi.fn(async()=>{});
 vi.stubGlobal("Excel",{run:async(fn:(c:unknown)=>Promise<unknown>)=>fn({workbook:{worksheets:{getItem:(id:string)=>({set position(position:number){write(id,position);}})}},sync})});
});
afterEach(()=>vi.unstubAllGlobals());
it.each([["b",0,["b","a","c"]],["a",3,["b","c","a"]],["c",1,["a","c","b"]]] as [string,number,string[]][])("%sを境界%iへ単一書込で移動",async(id,boundary,order)=>{
 const result=await reorderWorksheet(workbook(),id,boundary);expect(result.error).toBeUndefined();expect(result.state?.sheets.map(s=>s.id)).toEqual(order);expect(write).toHaveBeenCalledTimes(1);expect(result.state?.activeId).toBe("a");
});
it.each([1,2])("同位置境界%iには書き込まない",async boundary=>{const result=await reorderWorksheet(workbook(),"b",boundary);expect(result.error).toBeUndefined();expect(write).not.toHaveBeenCalled();});
it.each(["readOnly","structureProtected"] as const)("%sなら送信せず再取得",async key=>{const before=workbook();current[key]=true;const result=await reorderWorksheet(before,"b",0);expect(result.error).toBeTruthy();expect(write).not.toHaveBeenCalled();expect(loadWorkbookState).toHaveBeenCalledTimes(2);});
it.each([null,4,2,NaN,3.5])("総シート数%sが不明/不一致なら拒否",async total=>{current.totalSheets=total;const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toContain("混在");expect(write).not.toHaveBeenCalled();});
it("位置欠番は拒否",async()=>{current.sheets[2].position=4;const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toContain("混在");expect(write).not.toHaveBeenCalled();});
it("対象消失を拒否",async()=>{current.sheets.pop();current.totalSheets=2;const result=await reorderWorksheet(workbook(),"c",0);expect(result.error).toBeTruthy();expect(write).not.toHaveBeenCalled();});
it("直前の並び順変更を拒否",async()=>{write("c",0);write.mockClear();const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toContain("変わりました");expect(write).not.toHaveBeenCalled();});
it("同IDの名前・表示状態変化を拒否",async()=>{current.sheets[1].name="changed";current.sheets[1].visibility="Hidden";const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toBeTruthy();expect(write).not.toHaveBeenCalled();});
it("ブック切替世代を拒否",async()=>{const result=await reorderWorksheet(workbook(),"b",0,()=>false);expect(result.error).toContain("変わりました");expect(write).not.toHaveBeenCalled();});
it("対象シート保護を拒否",async()=>{current.sheets[1].protected=true;const result=await reorderWorksheet(structuredClone(current),"b",0);expect(result.error).toContain("保護");expect(write).not.toHaveBeenCalled();});
it("Hidden/VeryHiddenを可視化せず移動",async()=>{current.sheets[1].visibility="VeryHidden";const result=await reorderWorksheet(structuredClone(current),"b",0);expect(result.state?.sheets[0].visibility).toBe("VeryHidden");expect(result.error).toBeUndefined();});
it("書込API失敗は状態を再取得しエラーを返す",async()=>{sync.mockRejectedValueOnce(new Error("offline"));const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toContain("適用された可能性");expect(result.state?.sheets[0].id).toBe("b");expect(loadWorkbookState).toHaveBeenCalledTimes(2);});
it("書込後の確認失敗と再取得失敗は状態不明",async()=>{vi.mocked(loadWorkbookState).mockResolvedValueOnce(workbook()).mockRejectedValue(new Error("offline"));const result=await reorderWorksheet(workbook(),"b",0);expect(result.state).toBeUndefined();expect(result.error).toContain("再取得にも失敗");});
it("返る順序が期待と違えば失敗",async()=>{write.mockImplementation(()=>{});const result=await reorderWorksheet(workbook(),"b",0);expect(result.error).toContain("期待と一致");});
it("アクティブ状態が不明なら拒否",()=>{current.activeId=null;expect(moveRestriction(current)).toContain("アクティブ");});

it.each([["d",1,["a","d","b","c","e"]],["b",3,["a","c","b","d","e"]],["b",4,["a","c","d","b","e"]]] as [string,number,string[]][])("検索UIの全体境界で%sを移動し対象外を保持",async(id,gap,order)=>{
 current.sheets=["a","b","c","d","e"].map((id,position)=>({...current.sheets[0],id,name:id,position,isActive:id==="a"}));current.totalSheets=5;
 const result=await reorderWorksheet(structuredClone(current),id,gap);expect(result.error).toBeUndefined();expect(result.state?.sheets.map(s=>s.id)).toEqual(order);expect(write).toHaveBeenCalledOnce();
});

function displayBaseline(){const state=workbook();state.totalSheets=null;state.readPurpose="display";return state;}
it("表示用の総数未確認baselineからも厳密な確認成功後だけ移動する",async()=>{
 const baseline=displayBaseline();expect(dragRestriction(baseline)).toBeNull();expect(moveRestriction(baseline)).not.toBeNull();
 const result=await reorderWorksheet(baseline,"b",0);expect(result.error).toBeUndefined();expect(write).toHaveBeenCalledExactlyOnceWith("b",0);expect(loadWorkbookState).toHaveBeenCalledTimes(2);expect(result.state?.totalSheets).toBe(3);
});
it.each([null,4,2,NaN])("表示からの移動でも直前の総数%sなら変更を禁止",async total=>{
 current.totalSheets=total;const result=await reorderWorksheet(displayBaseline(),"b",0);expect(result.error).toContain("混在");expect(write).not.toHaveBeenCalled();
});
it.each(["readOnly","structureProtected"] as const)("表示からの移動でも直前の%sを拒否",async key=>{
 current[key]=true;const result=await reorderWorksheet(displayBaseline(),"b",0);expect(result.error).toBeTruthy();expect(write).not.toHaveBeenCalled();
});
it.each(["name","position","visibility","protected"] as const)("表示baselineの対象%s変更を検出して中止",async key=>{
 const baseline=displayBaseline();if(key==="name")current.sheets[1].name="renamed";
 if(key==="position"){write("b",0);write.mockClear();}if(key==="visibility")current.sheets[1].visibility="Hidden";if(key==="protected")current.sheets[1].protected=true;
 const result=await reorderWorksheet(baseline,"b",0);expect(result.error).toBeTruthy();expect(write).not.toHaveBeenCalled();
});
it("表示baselineからの移動後に総数不明なら成功扱いせず再取得する",async()=>{
 vi.mocked(loadWorkbookState).mockResolvedValueOnce(workbook()).mockImplementation(async()=>({...structuredClone(current),totalSheets:null}));
 const result=await reorderWorksheet(displayBaseline(),"b",0);expect(write).toHaveBeenCalledOnce();expect(result.error).toContain("移動結果を確認できません");expect(result.state?.totalSheets).toBeNull();expect(loadWorkbookState).toHaveBeenCalledTimes(3);
});
it("総数確認を試みて失敗した状態は表示用と混同せずドラッグを禁止",()=>{
 const unknown=workbook();unknown.totalSheets=null;expect(dragRestriction(unknown)).not.toBeNull();
 const mixed=displayBaseline();mixed.sheets[1].position=4;expect(dragRestriction(mixed)).not.toBeNull();
});
