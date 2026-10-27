import { afterEach, expect, it, vi } from "vitest";
import { CoalescedRefresh, SubscriptionLifecycle } from "../src/sync/coalescedRefresh";
import { withExcelLock } from "../src/excel/serial";
afterEach(()=>vi.useRealTimers());
it("イベントの連打を一回の取得へ集約",async()=>{
  vi.useFakeTimers();const read=vi.fn(async()=>{});const q=new CoalescedRefresh(read,vi.fn());
  for(let i=0;i<100;i++)q.request();await vi.advanceTimersByTimeAsync(150);expect(read).toHaveBeenCalledTimes(1);q.stop();
});
it("取得中の複数イベントを次の一回へ持ち越し重ねて実行しない",async()=>{
  vi.useFakeTimers();let release!:()=>void;const read=vi.fn(()=>new Promise<void>(resolve=>{release=resolve;}));
  const q=new CoalescedRefresh(read,vi.fn());q.request();await vi.advanceTimersByTimeAsync(150);
  q.request();q.request();await vi.advanceTimersByTimeAsync(500);expect(read).toHaveBeenCalledTimes(1);
  release();await vi.advanceTimersByTimeAsync(150);expect(read).toHaveBeenCalledTimes(2);release();q.stop();
});
it("停止後は予約・新規更新を実行しない",async()=>{vi.useFakeTimers();const read=vi.fn(async()=>{});const q=new CoalescedRefresh(read,vi.fn());q.request();q.stop();q.request();await vi.runAllTimersAsync();expect(read).not.toHaveBeenCalled();});
it("同期失敗を通知し、後続イベントで再試行できる",async()=>{vi.useFakeTimers();const read=vi.fn(async()=>{}).mockRejectedValueOnce(new Error("offline"));const error=vi.fn();const q=new CoalescedRefresh(read,error);q.request();await vi.advanceTimersByTimeAsync(150);expect(error).toHaveBeenCalledOnce();q.request();await vi.advanceTimersByTimeAsync(150);expect(read).toHaveBeenCalledTimes(2);q.stop();});
it("同時の有効化でも登録は1回",async()=>{const dispose=vi.fn(async()=>{}),register=vi.fn(async()=>dispose);const life=new SubscriptionLifecycle(register);await Promise.all([life.setEnabled(true),life.setEnabled(true)]);expect(register).toHaveBeenCalledOnce();await life.setEnabled(false);expect(dispose).toHaveBeenCalledOnce();});
it("開始中の停止要求を登録後に処理し、再開時は再登録",async()=>{const dispose=vi.fn(async()=>{}),register=vi.fn(async()=>dispose);const life=new SubscriptionLifecycle(register);await Promise.all([life.setEnabled(true),life.setEnabled(false)]);expect(dispose).toHaveBeenCalledOnce();await life.setEnabled(true);expect(register).toHaveBeenCalledTimes(2);await life.setEnabled(false);});
it("解除失敗は状態不明として再登録禁止",async()=>{const register=vi.fn(async()=>async()=>{throw new Error("offline");});const life=new SubscriptionLifecycle(register);await life.setEnabled(true);await expect(life.setEnabled(false)).rejects.toThrow("offline");await expect(life.setEnabled(true)).rejects.toThrow("状態が不明");expect(register).toHaveBeenCalledOnce();});
it("登録途中失敗でも再登録しない",async()=>{const register=vi.fn(async():Promise<()=>Promise<void>>=>{throw new Error("partial");});const life=new SubscriptionLifecycle(register);await expect(life.setEnabled(true)).rejects.toThrow();await expect(life.setEnabled(true)).rejects.toThrow("状態が不明");expect(register).toHaveBeenCalledOnce();});
it("手動更新・自動更新・変更操作を直列化し、失敗後も処理できる",async()=>{
  const order:string[]=[];let release!:()=>void;
  const first=withExcelLock(async()=>{order.push("start");await new Promise<void>(r=>{release=r;});order.push("end");});
  const second=withExcelLock(async()=>{order.push("next");throw new Error("fail");});const caught=second.catch(()=>undefined);
  await Promise.resolve();expect(order).toEqual(["start"]);release();await first;await caught;
  await withExcelLock(async()=>{order.push("last");});expect(order).toEqual(["start","end","next","last"]);
});
