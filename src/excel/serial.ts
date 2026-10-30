import { diagnose, measured } from "../diagnostics";
// 同一タスクペイン内のAPI処理を直列化。Excel標準UIや他クライアントは排他できない。
let tail: Promise<unknown> = Promise.resolve();
export function withExcelLock<T>(work: () => Promise<T>): Promise<T> {
  const queued=performance.now();diagnose("api.queue.request");
  const result = tail.then(()=>{diagnose("api.queue.wait",undefined,Math.round(performance.now()-queued));return measured("api.lock",work);});
  tail = result.catch(() => undefined);
  return result;
}
