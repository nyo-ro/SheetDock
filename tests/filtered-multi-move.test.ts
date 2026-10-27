import { expect, it } from "vitest";
import { groupOrder, positionPlan } from "../src/domain/movePlan";
import { runExperiment } from "../src/poc/experiment";
import { book, memoryAdapter } from "./poc-fixtures";
const full=["a","b","c","d","e"];
it.each([
 ["d",0,["a","d","b","c","e"]],
 ["b",1,["a","c","b","d","e"]],
 ["b",2,["a","c","d","b","e"]],
 ["b",0,full], ["d",2,full]
] as [string,number,string[]][])("検索対象b,dの%sを境界%iへ",(id,boundary,expected)=>expect(groupOrder(full,["b","d"],[id],boundary)).toEqual(expected));
function plan(selected:string[],boundary:number){const expectedOrder=groupOrder(full,full,selected,boundary);return {expectedOrder,actions:positionPlan(full,selected,expectedOrder).map(write=>({kind:"move" as const,...write}))};}
it.each([
 [["b","c"],0,["b","c","a","d","e"]],
 [["b","d"],0,["b","d","a","c","e"]],
 [["e","b"],3,["a","c","b","e","d"]],
 [["b","d"],5,["a","c","e","b","d"]],
 [["b","c"],2,full]
] as [string[],number,string[]][])("複数%sを%iへ実行して実際の順序を照合",async(selected,boundary,order)=>{
 const initial=book(),adapter=memoryAdapter(initial),p=plan(selected,boundary);
 const result=await runExperiment(adapter,initial,p.actions,{worksheetsOnly:true,expectedOrder:p.expectedOrder});
 expect(result.actualOrder).toEqual(order);expect(result.orderMatches).toBe(true);expect(result.steps.every(step=>step.status==="success")).toBe(true);
 if(order.join()===full.join())expect(adapter.apply).not.toHaveBeenCalled();
});
it.each(["before","after"] as const)("2番目の%s合成失敗で中断して再取得",async when=>{
 const initial=book(),adapter=memoryAdapter(initial),p=plan(["b","d"],0);
 const result=await runExperiment(adapter,initial,p.actions,{worksheetsOnly:true,expectedOrder:p.expectedOrder,injectFailure:{step:2,when}});
 expect(result.steps[0].status).toBe("success");expect(result.steps[1].status).toBe(when==="before"?"failed":"unknown");expect(result.steps.slice(2).every(step=>step.status==="unprocessed")).toBe(true);
 expect(adapter.apply).toHaveBeenCalledTimes(when==="before"?1:2);expect(result.after).toEqual(await adapter.read());expect(result.orderMatches).toBe(false);
});
it.each([null,undefined,6,4,NaN])("自己申告済みでも総数%sでは複数移動禁止",async total=>{
 const initial=book();initial.totalSheets=total;const adapter=memoryAdapter(initial),p=plan(["b","d"],0);
 const result=await runExperiment(adapter,initial,p.actions,{worksheetsOnly:true,expectedOrder:p.expectedOrder});
 expect(adapter.apply).not.toHaveBeenCalled();expect(result.steps[0].status).toBe("failed");expect(result.steps.slice(1).every(step=>step.status==="unprocessed")).toBe(true);
});
it("途中で未対応シートが追加されたら以後の書込を止める",async()=>{
 const initial=book(),adapter=memoryAdapter(initial),p=plan(["b","d"],0);const perform=adapter.apply.getMockImplementation()!;
 adapter.apply.mockImplementation(async action=>{await perform(action);adapter.change(b=>{b.totalSheets=6;});});
 const result=await runExperiment(adapter,initial,p.actions,{worksheetsOnly:true,expectedOrder:p.expectedOrder});
 expect(adapter.apply).toHaveBeenCalledOnce();expect(result.steps[0].status).toBe("unknown");expect(result.steps[1].status).toBe("unprocessed");
});

it("2回目のAPI呼出しが例外なら状態不明として後続を止める",async()=>{
 const initial=book(),adapter=memoryAdapter(initial),p=plan(["b","d"],0);const perform=adapter.apply.getMockImplementation()!;let count=0;
 adapter.apply.mockImplementation(async action=>{if(++count===2)throw new Error("AccessDenied mock");await perform(action);});
 const result=await runExperiment(adapter,initial,p.actions,{worksheetsOnly:true,expectedOrder:p.expectedOrder});
 expect(result.steps.map(s=>s.status)).toEqual(["success","unknown",...p.actions.slice(2).map(()=>"unprocessed")]);expect(adapter.apply).toHaveBeenCalledTimes(2);expect(result.after).toEqual(await adapter.read());
});
it("同位置計画でも混在不明なら再検証で拒否",async()=>{const initial=book();initial.totalSheets=null;const adapter=memoryAdapter(initial);await expect(runExperiment(adapter,initial,[],{worksheetsOnly:true,expectedOrder:full})).rejects.toThrow("混在");expect(adapter.apply).not.toHaveBeenCalled();});
