import { expect, it } from "vitest";
import { groupOrder, positionPlan } from "../src/domain/movePlan";
import { runExperiment } from "../src/poc/experiment";
import { Action } from "../src/poc/model";
import { book, memoryAdapter } from "./poc-fixtures";
const options={worksheetsOnly:true};
it.each([
  ["Visible","Hidden"],["Visible","VeryHidden"],["Hidden","Visible"],["VeryHidden","Visible"],["VeryHidden","Hidden"]
] as const)("%s→%sを再取得で確認",async (from,to) => {
  const b=book();b.sheets[1].visibility=from;const adapter=memoryAdapter(b);
  const r=await runExperiment(adapter,b,[{kind:"visibility",id:"b",visibility:to}],options);
  expect(r.steps[0].status).toBe("success");expect(r.after?.sheets[1].visibility).toBe(to);
});
it.each(["Hidden","VeryHidden"] as const)("%s複製は元状態を変更せず複製先を観測",async visibility => {
  const b=book();b.sheets[1].visibility=visibility;const adapter=memoryAdapter(b);
  const r=await runExperiment(adapter,b,[{kind:"copy",id:"b"}],options);
  expect(r.steps[0].status).toBe("success");expect(adapter.apply).toHaveBeenCalledTimes(1);
  expect(r.after?.sheets.find(s => s.id === "b")?.visibility).toBe(visibility);
});
it("Hidden削除を確認",async () => {const b=book();b.sheets[1].visibility="Hidden";const r=await runExperiment(memoryAdapter(b),b,[{kind:"delete",id:"b"}],options);expect(r.steps[0].status).toBe("success");expect(r.after?.sheets).toHaveLength(4);});
it("非表示混在の複数移動を計画順に実行し最終順序を照合",async () => {
  const b=book();b.sheets[1].visibility="Hidden";b.sheets[3].visibility="VeryHidden";
  const full=b.sheets.map(s=>s.id), selected=["b","d"],desired=groupOrder(full,full,selected,0);
  const actions: Action[]=positionPlan(full,selected,desired).map(s=>({kind:"move",...s}));
  const r=await runExperiment(memoryAdapter(b),b,actions,{...options,expectedOrder:desired});
  expect(r.steps.every(s=>s.status === "success")).toBe(true);expect(r.orderMatches).toBe(true);
});
it("直前の名前変更ではAPIを呼ばない",async () => {
  const b=book(),a=memoryAdapter(b);a.change(s=>{s.sheets[1].name="changed";});
  const r=await runExperiment(a,b,[{kind:"copy",id:"b"}],options);
  expect(r.steps[0].status).toBe("failed");expect(a.apply).not.toHaveBeenCalled();
});
it("成功・失敗・未処理を分離し最初の失敗で停止",async () => {
  const b=book(),a=memoryAdapter(b); const actions:Action[]=["b","c","d"].map(id=>({kind:"visibility",id,visibility:"Hidden"}));
  const r=await runExperiment(a,b,actions,{...options,injectFailure:{step:2,when:"before"}});
  expect(r.steps.map(s=>s.status)).toEqual(["success","failed","unprocessed"]);expect(a.apply).toHaveBeenCalledTimes(1);
});
it("書込後エラーは状態不明。再取得結果があっても成功に変換しない",async () => {
  const b=book(),a=memoryAdapter(b); const actions:Action[]=["b","c"].map(id=>({kind:"visibility",id,visibility:"Hidden"}));
  const r=await runExperiment(a,b,actions,{...options,injectFailure:{step:1,when:"after"}});
  expect(r.steps.map(s=>s.status)).toEqual(["unknown","unprocessed"]);expect(r.after?.sheets[1].visibility).toBe("Hidden");
});
it("実APIを模したrejectでも後続を停止",async () => {
  const b=book(),a=memoryAdapter(b);a.apply.mockRejectedValueOnce(new Error("AccessDenied"));
  const r=await runExperiment(a,b,[{kind:"copy",id:"b"},{kind:"copy",id:"c"}],options);
  expect(r.steps.map(s=>s.status)).toEqual(["unknown","unprocessed"]);
});
it("API成功でも結果不一致は状態不明",async () => {
  const b=book(),a=memoryAdapter(b);a.apply.mockResolvedValueOnce(undefined);
  const r=await runExperiment(a,b,[{kind:"delete",id:"b"}],options);expect(r.steps[0].status).toBe("unknown");
});
const vhActions:Action[]=[{kind:"visibility",id:"b",visibility:"Hidden"},{kind:"delete",id:"b"}];
it("VeryHidden削除を2段階で観測",async () => {
  const b=book();b.sheets[1].visibility="VeryHidden";
  const r=await runExperiment(memoryAdapter(b),b,vhActions,{...options,veryHiddenDeleteId:"b"});
  expect(r.steps.map(s=>s.status)).toEqual(["success","success"]);expect(r.safeForProduction).toBe(false);
});
it("削除前の失敗時にVeryHidden復元を試行・再確認",async () => {
  const b=book();b.sheets[1].visibility="VeryHidden";
  const r=await runExperiment(memoryAdapter(b),b,vhActions,{...options,veryHiddenDeleteId:"b",injectFailure:{step:2,when:"before"}});
  expect(r.recovery?.status).toBe("restored");expect(r.after?.sheets[1].visibility).toBe("VeryHidden");
});
it("復元APIも失敗したら復元保証をしない",async () => {
  const b=book();b.sheets[1].visibility="VeryHidden";const a=memoryAdapter(b);const original=a.apply.getMockImplementation()!;
  a.apply.mockImplementation(async action=>{if(action.kind === "visibility" && action.visibility === "VeryHidden") throw new Error("offline");await original(action);});
  const r=await runExperiment(a,b,vhActions,{...options,veryHiddenDeleteId:"b",injectFailure:{step:2,when:"before"}});
  expect(r.recovery?.status).toBe("unknown");expect(r.after).toBeUndefined();
});
it("復元前に同じIDの名前が変われば復元を書き込まない",async () => {
  const b=book();b.sheets[1].visibility="VeryHidden";const a=memoryAdapter(b);const original=a.apply.getMockImplementation()!;
  a.apply.mockImplementation(async action=>{if(action.kind === "delete") {a.change(s=>{s.sheets[1].name="replacement";});throw new Error("conflict");}await original(action);});
  const r=await runExperiment(a,b,vhActions,{...options,veryHiddenDeleteId:"b"});
  expect(r.recovery?.status).toBe("unknown");expect(a.apply).toHaveBeenCalledTimes(2);
});
it("一時変更自体が不明なら安全な復元とは扱わない",async () => {
  const b=book();b.sheets[1].visibility="VeryHidden";
  const r=await runExperiment(memoryAdapter(b),b,vhActions,{...options,veryHiddenDeleteId:"b",injectFailure:{step:1,when:"after"}});
  expect(r.recovery?.status).toBe("unknown");expect(r.steps[1].status).toBe("unprocessed");
});
it("一括途中の外部順序変更を再検証で止める",async () => {
  const b=book(),a=memoryAdapter(b);const read=a.read.getMockImplementation()!;let count=0;
  a.read.mockImplementation(async()=>{count++;if(count===3)a.change(s=>{s.sheets[0].position=4;});return read();});
  const r=await runExperiment(a,b,[{kind:"visibility",id:"b",visibility:"Hidden"},{kind:"visibility",id:"c",visibility:"Hidden"}],options);
  expect(r.steps.map(s=>s.status)).toEqual(["success","failed"]);expect(a.apply).toHaveBeenCalledTimes(1);
});

it("書込後のアクティブ取得だけが失敗しても状態不明にする",async () => {
  const b=book(),a=memoryAdapter(b);const original=a.apply.getMockImplementation()!;
  a.apply.mockImplementation(async action=>{await original(action);a.change(state=>{state.activeId=null;state.activeError="active read failed";});});
  const r=await runExperiment(a,b,[{kind:"visibility",id:"b",visibility:"Hidden"}],options);
  expect(r.steps[0].status).toBe("unknown");expect(r.after?.activeError).toBe("active read failed");
});
