import { SubscriptionLifecycle } from "../sync/coalescedRefresh";
import { withExcelLock } from "../excel/serial";
export const eventSets = {
  onActivated:"1.7", onAdded:"1.7", onDeleted:"1.7", onNameChanged:"1.17",
  onMoved:"1.17", onVisibilityChanged:"1.17", onProtectionChanged:"1.14"
} as const;
export interface EventRecord { time: string; type: string; args: unknown }
const listeners = new Set<(record: EventRecord) => void>();
export function listenEvents(listener: (record: EventRecord) => void): () => void { listeners.add(listener); return () => { listeners.delete(listener); }; }
export async function registerWorksheetEvents(emit: (record: EventRecord) => void): Promise<() => Promise<void>> {
  return withExcelLock(() => Excel.run(async context => {
    const handles: {remove(): void}[] = [];
    try {
      for (const [type, set] of Object.entries(eventSets)) {
        if (!Office.context.requirements.isSetSupported("ExcelApi",set)) throw new Error(`${type}: ExcelApi ${set}が必要です。`);
        const event = context.workbook.worksheets[type as keyof typeof eventSets];
        handles.push(event.add(async (args: unknown) => emit({time:new Date().toISOString(),type,args})));
      }
      await context.sync();
    } catch (error) {
      // 登録途中の失敗でも解除を試行。Lifecycleは状態不明として再登録を禁止。
      try { handles.forEach(h => h.remove()); await context.sync(); } catch { /* ペイン再読込が必要 */ }
      throw error;
    }
    return () => withExcelLock(() => Excel.run(context, async sameContext => {
      handles.forEach(h => h.remove()); await sameContext.sync();
    }));
  }));
}
export const eventLifecycle = new SubscriptionLifecycle(() => registerWorksheetEvents(record => listeners.forEach(l => l(record))));
