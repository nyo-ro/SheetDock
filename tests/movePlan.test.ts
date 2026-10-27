import { describe, expect, it } from "vitest";
import { groupOrder, positionPlan } from "../src/domain/movePlan";
const full = ["a","b","c","d","e"];
describe("Sheaside互換の境界移動", () => {
  it.each([
    [["b","c"],0,["b","c","a","d","e"]],
    [["b","d"],5,["a","c","e","b","d"]],
    [["b","d"],0,["b","d","a","c","e"]],
    [["b","c"],2,full],
    [["e"],0,["e","a","b","c","d"]],
    [["a"],5,["b","c","d","e","a"]],
    [["d"],1,["a","d","b","c","e"]],
    [["b"],4,["a","c","d","b","e"]],
    [["d","b"],3,["a","c","b","d","e"]]
  ] as [string[],number,string[]][])("%j を境界 %i へ",(selected,boundary,expected) => expect(groupOrder(full,full,selected,boundary)).toEqual(expected));
  it("絞り込み末尾は最後の表示対象の直後", () => expect(groupOrder(full,["a","c"],["a"],2)).toEqual(["b","c","a","d","e"]));
  it.each([-1,6,0.5,NaN])("不正境界 %s を拒否", boundary => expect(() => groupOrder(full,full,["a"],boundary)).toThrow());
  it("重複・対象消失・表示順の不整合を拒否", () => {
    expect(() => groupOrder(full,full,["a","a"],0)).toThrow();
    expect(() => groupOrder(full,["x"],["x"],0)).toThrow();
    expect(() => groupOrder(full,["c","a"],["a"],0)).toThrow();
  });
  it("全非空部分集合×全境界で、選択だけを書き換え最終順序を再現する", () => {
    for (let mask = 1; mask < (1 << full.length); mask++) {
      const selected = full.filter((_,i) => mask & (1 << i));
      for (let boundary = 0; boundary <= full.length; boundary++) {
        const desired = groupOrder(full,full,selected,boundary);
        const plan = positionPlan(full,selected,desired); const actual = [...full];
        for (const step of plan) {
          expect(selected).toContain(step.id);
          actual.splice(step.position,0,actual.splice(actual.indexOf(step.id),1)[0]);
          expect(actual).toEqual(step.orderAfter);
        }
        expect(actual).toEqual(desired);
        expect(actual.filter(id => selected.includes(id))).toEqual(selected);
        expect(actual.filter(id => !selected.includes(id))).toEqual(full.filter(id => !selected.includes(id)));
      }
    }
  });
  it("同位置は書き込みなし", () => expect(positionPlan(full,["b"],full)).toEqual([]));
  it("移動対象内の相対順序変更を拒否", () => expect(() => positionPlan(full,["b","d"],["a","d","b","c","e"])).toThrow());
});
