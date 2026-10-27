import { expect, it, vi } from "vitest";
import { SharedEvents } from "../src/sync/sharedEvents";
it("通常UIとPoCの同時購読でも1登録、最後の解除のみ物理解除",async()=>{
 let emit!:(event:string)=>void;const dispose=vi.fn(async()=>{});const register=vi.fn(async(callback:(event:string)=>void)=>{emit=callback;return dispose;});
 const hub=new SharedEvents(register);const a=vi.fn(),b=vi.fn();const [releaseA,releaseB]=await Promise.all([hub.subscribe(a),hub.subscribe(b)]);
 expect(register).toHaveBeenCalledTimes(1);emit("changed");expect(a).toHaveBeenCalledOnce();expect(b).toHaveBeenCalledOnce();
 await releaseA();expect(dispose).not.toHaveBeenCalled();emit("again");expect(a).toHaveBeenCalledTimes(1);expect(b).toHaveBeenCalledTimes(2);
 await releaseB();await releaseB();expect(dispose).toHaveBeenCalledTimes(1);
});
it("ブック再登録は旧解除→新登録、購読者は維持",async()=>{const order:string[]=[];const register=vi.fn(async()=>{order.push("add");return async()=>{order.push("remove");};});const hub=new SharedEvents(register);const release=await hub.subscribe(vi.fn());await hub.rebind();expect(order).toEqual(["add","remove","add"]);await release();});
it("解除済みと確認できる登録失敗なら再試行できる",async()=>{const register=vi.fn(async()=>async()=>{}).mockRejectedValueOnce(new Error("offline"));const hub=new SharedEvents(register);await expect(hub.subscribe(vi.fn())).rejects.toThrow();const release=await hub.subscribe(vi.fn());expect(register).toHaveBeenCalledTimes(2);await release();});
it("登録状態不明なら重複を避け再登録を拒否",async()=>{const register=vi.fn(async()=>{throw Object.assign(new Error("unknown"),{uncertain:true});});const hub=new SharedEvents(register);await expect(hub.subscribe(vi.fn())).rejects.toThrow();await expect(hub.subscribe(vi.fn())).rejects.toThrow("状態が不明");expect(register).toHaveBeenCalledOnce();});
it("解除失敗後は再登録しない",async()=>{const register=vi.fn(async()=>async()=>{throw new Error("offline");});const hub=new SharedEvents(register);const release=await hub.subscribe(vi.fn());await expect(release()).rejects.toThrow();await expect(hub.subscribe(vi.fn())).rejects.toThrow("状態が不明");expect(register).toHaveBeenCalledOnce();});

it("再登録に一度失敗しても既存購読を保って再試行",async()=>{let emit!:(event:string)=>void;const listener=vi.fn();const register=vi.fn(async(callback:(event:string)=>void)=>{emit=callback;return async()=>{};});const hub=new SharedEvents(register);const release=await hub.subscribe(listener);register.mockRejectedValueOnce(new Error("offline"));await expect(hub.rebind()).rejects.toThrow("offline");await hub.rebind();emit("recovered");expect(listener).toHaveBeenCalledWith("recovered");expect(register).toHaveBeenCalledTimes(3);await release();});
