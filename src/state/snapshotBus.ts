import { SheetInfo } from "../domain/sheets";
export type SnapshotNotice = { sheets: SheetInfo[]; error?: string };
const listeners = new Set<(notice: SnapshotNotice) => void>();
export function publishSnapshot(notice: SnapshotNotice): void { listeners.forEach(listener => listener(notice)); }
export function subscribeSnapshot(listener: (notice: SnapshotNotice) => void): () => void {
  listeners.add(listener); return () => { listeners.delete(listener); };
}
