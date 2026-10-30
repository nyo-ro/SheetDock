// 開発専用の読み取り比較。呼出側が1実行/ページと排他を管理する。
// 通常UIのキュー・イベント・取得処理には依存しない。
export type ProbeKind = "A" | "B" | "C";
export interface SheetChoice { id: string; name: string; position: number }
export interface ProbeResult { count: number; choices?: SheetChoice[] }
export type ProbeTrace = (stage: string, error?: unknown, durationMs?: number) => void;

async function measured<T>(stage: string, work: () => Promise<T>, trace: ProbeTrace): Promise<T> {
  const started = performance.now();
  trace(stage + ".start");
  try { return await work(); }
  catch (error) { trace(stage + ".error", error); throw error; }
  finally { trace(stage + ".end", undefined, Math.round(performance.now() - started)); }
}

export async function runReadProbe(kind: ProbeKind, knownId: string | undefined, trace: ProbeTrace): Promise<ProbeResult> {
  if (kind !== "A" && kind !== "B" && kind !== "C") {
    throw Object.assign(new Error("未知の実験種別です。"), { code: "ProbeInvalidKind" });
  }
  if (kind === "B" && (typeof knownId !== "string" || !knownId.trim())) {
    throw Object.assign(new Error("準備で取得した対象シートを選んでください。"), { code: "ProbeTargetRequired" });
  }
  return measured("api.run", () => Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    if (kind === "A") {
      sheets.load("items/id,items/name,items/position");
      await measured("api.sync", () => context.sync(), trace);
      const choices = sheets.items.map(sheet => ({ id: sheet.id, name: sheet.name, position: sheet.position }));
      return { count: choices.length, choices };
    }
    if (kind === "B") {
      const target = sheets.getItemOrNullObject(knownId!);
      target.load("id,name");
      await measured("api.sync", () => context.sync(), trace);
      if (target.isNullObject) throw Object.assign(new Error("対象のシートが見つかりません。この試験を終了してください。"), { code: "ProbeTargetMissing" });
      // 値を読むが、ID・名前は診断結果へ含めない。
      if (typeof target.id !== "string" || typeof target.name !== "string") {
        throw Object.assign(new Error("対象シートの情報を取得できませんでした。"), { code: "ProbeInvalidTarget" });
      }
      if (target.id !== knownId) throw Object.assign(new Error("取得したシートIDが指定したIDと一致しません。"), { code: "ProbeTargetMismatch" });
      return { count: 1 };
    }
    const active = sheets.getActiveWorksheet();
    active.load("id");
    await measured("api.sync", () => context.sync(), trace);
    if (typeof active.id !== "string" || !active.id) throw Object.assign(new Error("アクティブシートIDを取得できませんでした。"), { code: "ProbeActiveIdMissing" });
    return { count: 1 };
  }), trace);
}

// Bの準備は本試験と別の段階として呼び出し、traceの接頭辞は呼出側が付ける。
export async function prepareKnownSheets(trace: ProbeTrace): Promise<SheetChoice[]> {
  const result = await runReadProbe("A", undefined, trace);
  return result.choices!;
}
