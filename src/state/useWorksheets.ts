import { useCallback, useEffect, useRef, useState } from "react";
import { SheetInfo } from "../domain/sheets";
import { ensureExcelReady, errorMessage } from "../excel/environment";
import { activateWorksheet, readWorksheets } from "../excel/worksheets";

export function useWorksheets() {
  const [sheets, setSheets] = useState<SheetInfo[]>([]);
  const [query, setQuery] = useState("");
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [hasSnapshot, setHasSnapshot] = useState(false);
  const locked = useRef(false);
  const mounted = useRef(false);
  useEffect(() => {
    mounted.current = true;
    let cancelled = false;
    void (async () => {
      try {
        await ensureExcelReady();
        if (cancelled) return;
        setReady(true);
        const result = await readWorksheets();
        if (!cancelled) { setSheets(result); setHasSnapshot(true); }
      } catch (failure) {
        if (!cancelled) setError(errorMessage(failure));
      } finally {
        if (!cancelled) setBusy(false);
      }
    })();
    return () => { cancelled = true; mounted.current = false; };
  }, []);
  const run = useCallback(async (id?: string) => {
    if (!ready || busy || locked.current) return;
    locked.current = true;
    setBusy(true);
    setError(null);
    try {
      if (id !== undefined) await activateWorksheet(id);
      const result = await readWorksheets();
      if (mounted.current) { setSheets(result); setHasSnapshot(true); }
    } catch (failure) {
      // 途中失敗時は古いアクティブ表示を残さず、再取得を促す。
      if (mounted.current) { setSheets([]); setHasSnapshot(false); setError(errorMessage(failure)); }
    } finally {
      locked.current = false;
      if (mounted.current) setBusy(false);
    }
  }, [ready, busy]);
  return { sheets, query, setQuery, ready, busy, error, hasSnapshot,
    refresh: () => run(), activate: (id: string) => run(id) };
}
