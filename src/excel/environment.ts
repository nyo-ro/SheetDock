export async function ensureExcelReady(timeoutMs = 15000): Promise<void> {
  if (typeof Office === "undefined") {
    throw new Error("Office.jsを読み込めませんでした。ネットワークを確認し、Excelから開き直してください。");
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const info = await Promise.race([
      Office.onReady(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new Error("Excelとの接続がタイムアウトしました。Excelから開き直してください。")), timeoutMs);
      })
    ]);
    if (info.host !== Office.HostType.Excel) throw new Error("SheetDockはExcelのタスクペインから開いてください。");
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.17")) {
      throw new Error("このExcelはExcelApi 1.17に対応していません。対応するMicrosoft 365 Excelを使用してください。");
    }
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}
export function errorMessage(error: unknown): string {
  const detail = error instanceof Error ? error.message : "Excelとの通信中に問題が発生しました。";
  return detail;
}
