// 同一タスクペイン内のAPI処理を直列化。Excel標準UIや他クライアントは排他できない。
let tail: Promise<unknown> = Promise.resolve();
export function withExcelLock<T>(work: () => Promise<T>): Promise<T> {
  const result = tail.then(work);
  tail = result.catch(() => undefined);
  return result;
}
