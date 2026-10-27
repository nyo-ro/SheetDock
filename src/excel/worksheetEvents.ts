import { SharedEvents } from "../sync/sharedEvents";
import { withExcelLock } from "../excel/serial";
export const eventSets = {
  onActivated:"1.7", onAdded:"1.7", onDeleted:"1.7", onNameChanged:"1.17",
  onMoved:"1.17", onVisibilityChanged:"1.17", onProtectionChanged:"1.14"
} as const;
export interface EventRecord { time: string; type: string; args: unknown }
export async function registerWorksheetEvents(emit: (record: EventRecord) => void, workbookActivation = false): Promise<() => Promise<void>> {
  return withExcelLock(() => Excel.run(async context => {
    const handles: {remove(): void}[] = [];
    try {
      for (const [type, set] of Object.entries(eventSets)) {
        if (!Office.context.requirements.isSetSupported("ExcelApi",set)) throw new Error(`${type}: ExcelApi ${set}が必要です。`);
        const event = context.workbook.worksheets[type as keyof typeof eventSets];
        handles.push(event.add(async (args: unknown) => emit({time:new Date().toISOString(),type,args})));
      }
      if (workbookActivation) handles.push(context.workbook.onActivated.add(async args => emit({time:new Date().toISOString(),type:"workbookActivated",args})));
      await context.sync();
    } catch (error) {
      // 登録途中の失敗でも解除を試行。解除できなければ再登録を禁止。
      try { handles.forEach(h => h.remove()); await context.sync(); } catch { throw Object.assign(new Error("イベント登録状態が不明です。ペインを開き直してください。"),{uncertain:true}); }
      throw error;
    }
    return () => withExcelLock(() => Excel.run(context, async sameContext => {
      handles.forEach(h => h.remove()); await sameContext.sync();
    }));
  }));
}
export const worksheetEvents = new SharedEvents<EventRecord>(emit => registerWorksheetEvents(emit, true));
