import { expect, it } from "vitest";
import { assertAction, assertFresh, assertWritable } from "../src/poc/model";
import { book } from "./poc-fixtures";
it.each(["Hidden","VeryHidden"] as const)("最後の表示シートを%sにしない",visibility => {
  const b=book(); b.sheets.slice(1).forEach(s => {s.visibility="Hidden";});
  expect(() => assertAction(b,{kind:"visibility",id:"a",visibility})).toThrow("最後");
  expect(() => assertAction(b,{kind:"delete",id:"a"})).toThrow("最後");
});
it("最後の表示シートがあっても他の非表示シートは削除可能",() => {
  const b=book(); b.sheets[1].visibility="Hidden";
  expect(() => assertAction(b,{kind:"delete",id:"b"})).not.toThrow();
});
it("VeryHidden直接削除を止める",() => {const b=book(); b.sheets[1].visibility="VeryHidden";expect(() => assertAction(b,{kind:"delete",id:"b"})).toThrow("Hidden");});
it.each(["readOnly","structureProtected"] as const)("%sを拒否",property => {const b=book();b[property]=true;expect(() => assertWritable(b,true)).toThrow();});
it("混在・種別不明は変更不可",() => expect(() => assertWritable(book(),false)).toThrow("観測専用"));
it("位置の欠番から混在の疑いを検出",() => {const b=book();b.sheets[2].position=8;expect(() => assertWritable(b,true)).toThrow("欠番");});
it("アクティブ取得失敗は変更不可",() => {const b=book();b.activeId=null;expect(() => assertWritable(b,true)).toThrow("アクティブ");});
it("シート保護は保守的に拒否",() => {const b=book();b.sheets[0].protected=true;expect(() => assertAction(b,{kind:"copy",id:"a"})).toThrow("シート保護");});
it.each(["name","position","visibility","protected"])("同じIDでも%s変化を拒否", property => {
  const baseline=book(), actual=book();
  Object.assign(actual.sheets[0],{[property]:property === "name" ? "renamed" : property === "position" ? 9 : property === "visibility" ? "Hidden" : true});
  expect(() => assertFresh(baseline,actual)).toThrow("変化");
});
it("対象削除を検出",() => {const a=book();a.sheets.pop();expect(() => assertFresh(book(),a)).toThrow();});
it("取得時刻やアクティブ変更だけでは構成競合としない",() => {const a=book();a.capturedAt="later";a.activeId="b";expect(() => assertFresh(book(),a)).not.toThrow();});
