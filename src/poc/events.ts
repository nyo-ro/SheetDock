import { SubscriptionLifecycle } from "../sync/coalescedRefresh";
import { worksheetEvents, EventRecord } from "../excel/worksheetEvents";
export { registerWorksheetEvents, eventSets } from "../excel/worksheetEvents";
export type { EventRecord } from "../excel/worksheetEvents";
const listeners = new Set<(record:EventRecord)=>void>();
export function listenEvents(listener:(record:EventRecord)=>void):()=>void {listeners.add(listener);return ()=>{listeners.delete(listener);};}
export const eventLifecycle = new SubscriptionLifecycle(()=>worksheetEvents.subscribe(record=>listeners.forEach(l=>l(record))));
