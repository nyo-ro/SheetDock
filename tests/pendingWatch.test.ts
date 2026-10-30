import { afterEach, expect, it, vi } from "vitest";
import { PendingWatch } from "../src/sync/pendingWatch";
import { withExcelLock } from "../src/excel/serial";
afterEach(()=>vi.useRealTimers());
it("期限超過を通知してもOfficeロックを解放せず並列APIを開始しない",async()=>{
 vi.useFakeTimers();let finish!:()=>void;let valid!:()=>boolean;const changed=vi.fn(),settled=vi.fn(),next=vi.fn(async()=>{});const watch=new PendingWatch(changed,settled,100);
 const first=watch.run("test",ok=>{valid=ok;return withExcelLock(()=>new Promise<void>(r=>{finish=r;}));});await Promise.resolve();const second=withExcelLock(next);
 await vi.advanceTimersByTimeAsync(100);expect(changed).toHaveBeenCalledWith(true);expect(valid()).toBe(false);expect(next).not.toHaveBeenCalled();finish();await first;await second;expect(next).toHaveBeenCalledOnce();expect(settled).toHaveBeenCalledOnce();
});
it("エラーでも監視タイマーを解除し遅延通知しない",async()=>{vi.useFakeTimers();const changed=vi.fn(),settled=vi.fn();const watch=new PendingWatch(changed,settled,100);await expect(watch.run("test",async()=>{throw new Error("fail");})).rejects.toThrow("fail");await vi.advanceTimersByTimeAsync(100);expect(changed).not.toHaveBeenCalledWith(true);expect(settled).not.toHaveBeenCalled();});
it("複数処理の片方の終了で他方の長時間待機表示を解除しない",async()=>{vi.useFakeTimers();const changed=vi.fn();const watch=new PendingWatch(changed,vi.fn(),100);let first!:()=>void,second!:()=>void;const a=watch.run("a",()=>new Promise<void>(r=>{first=r;}));const b=watch.run("b",()=>new Promise<void>(r=>{second=r;}));await vi.advanceTimersByTimeAsync(100);first();await a;expect(changed).toHaveBeenLastCalledWith(true);second();await b;expect(changed).toHaveBeenLastCalledWith(false);});
