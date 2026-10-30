const modeKey = "sheetdock.readProbe";
export function isReadProbeMode(): boolean {
  try { return sessionStorage.getItem(modeKey) === "1"; } catch { return false; }
}
// モード変更は必ずページを再読み込みして適用。同じ実行環境に通常UIと実験を混在させない。
export function reloadInMode(experiment: boolean): void {
  if (experiment) sessionStorage.setItem(modeKey, "1"); else sessionStorage.removeItem(modeKey);
  window.location.reload();
}
