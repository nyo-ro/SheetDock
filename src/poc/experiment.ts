import { Action, BookSnapshot, ExperimentResult, assertAction, assertMoveWorkbook, assertFresh, assertWritable, describeError, fingerprint } from "./model";
export interface ExperimentAdapter { read(): Promise<BookSnapshot>; apply(action: Action): Promise<void> }
export interface ExperimentOptions {
  worksheetsOnly: boolean; expectedOrder?: string[]; veryHiddenDeleteId?: string;
  // PoC専用。実APIエラーと混同しないためログにSyntheticと明示する。
  injectFailure?: { step: number; when: "before" | "after" };
}
function verifyChange(before: BookSnapshot, after: BookSnapshot, action: Action): void {
  if (after.activeId === null || after.activeError) throw new Error("操作後のアクティブ状態を確認できません。状態不明として中断します。");
  const expected: BookSnapshot = structuredClone(before);
  const index = expected.sheets.findIndex(s => s.id === action.id);
  if (action.kind === "visibility") expected.sheets[index].visibility = action.visibility;
  if (action.kind === "move") expected.sheets = action.orderAfter.map((id, position) => {
    const found = before.sheets.find(s => s.id === id); if (!found) throw new Error("不正な移動先IDです。");
    return { ...found, position };
  });
  if (action.kind === "delete") expected.sheets = expected.sheets.filter(s => s.id !== action.id).map((s,position) => ({...s,position}));
  if (action.kind === "copy") {
    const added = after.sheets.filter(s => !before.sheets.some(old => old.id === s.id));
    if (added.length !== 1 || after.sheets.length !== before.sheets.length + 1) throw new Error("複製先を一意に確認できません。");
    // 複製先のvisibilityは推測せず観測値を記録。元シートの状態は厳密に確認。
    expected.sheets.splice(index + 1, 0, added[0]);
    expected.sheets = expected.sheets.map((s,position) => ({...s,position}));
  }
  if (fingerprint(expected) !== fingerprint(after)) throw new Error("操作後の状態が期待と一致しません。状態不明として中断します。");
}
export async function runExperiment(adapter: ExperimentAdapter, baseline: BookSnapshot, actions: Action[], options: ExperimentOptions): Promise<ExperimentResult> {
  const result: ExperimentResult = { before: baseline, steps: actions.map(action => ({action,status:"unprocessed"})), safeForProduction: false, expectedOrder: options.expectedOrder };
  const moving = !!options.expectedOrder || actions.some(action => action.kind === "move");
  let expected = baseline;
  let failed = false;
  for (let i = 0; i < actions.length; i++) {
    const step = result.steps[i]; let dispatched = false;
    try {
      const current = await adapter.read(); step.before = current;
      assertFresh(expected, current); assertWritable(current, options.worksheetsOnly); assertAction(current, step.action);
      if (moving) assertMoveWorkbook(current);
      if (options.injectFailure?.step === i + 1 && options.injectFailure.when === "before") throw new Error("SyntheticFailure: 書き込み前の検証用エラー");
      dispatched = true; await adapter.apply(step.action);
      if (options.injectFailure?.step === i + 1 && options.injectFailure.when === "after") throw new Error("SyntheticFailure: 書き込み後・結果確認前の検証用エラー");
      const after = await adapter.read(); step.after = after; result.after = after;
      if (moving) assertMoveWorkbook(after);
      verifyChange(current, after, step.action);
      expected = after; step.status = "success";
    } catch (error) {
      step.status = dispatched ? "unknown" : "failed"; step.error = describeError(error); failed = true;
      try { step.after = await adapter.read(); result.after = step.after; } catch (readError) {
        result.after = undefined; step.error = { operation: step.error, observation: describeError(readError) };
      }
      break;
    }
  }
  if (!actions.length) {
    result.after = await adapter.read(); assertFresh(baseline, result.after); assertWritable(result.after, options.worksheetsOnly); if (moving) assertMoveWorkbook(result.after);
  }
  // ベストエフォート復元。全例での保証はないので本番採用不可。
  if (failed && options.veryHiddenDeleteId) {
    const id = options.veryHiddenDeleteId;
    const original = baseline.sheets.find(s => s.id === id);
    const visibilityStep = result.steps.find(s => s.action.kind === "visibility" && s.action.id === id);
    // 復元対象を説明できるのは一時変更の成功を確認できた場合だけ。
    if (original?.visibility === "VeryHidden" && visibilityStep?.status === "success" && visibilityStep.after) {
      try {
        const current = await adapter.read();
        if (!current.sheets.some(s => s.id === id)) result.recovery = { status: "not-needed", after: current };
        else {
          assertFresh(visibilityStep.after, current); assertWritable(current, options.worksheetsOnly);
          const restore: Action = {kind:"visibility",id,visibility:"VeryHidden"}; assertAction(current, restore);
          await adapter.apply(restore); const after = await adapter.read(); verifyChange(current, after, restore);
          result.after = after; result.recovery = {status:"restored",after};
        }
      } catch (error) { result.recovery = {status:"unknown",error:describeError(error)}; result.after = undefined; }
    } else result.recovery = {status:"unknown",error:{message:"一時変更の成功が未確認のため、安全に復元対象を特定できません。"}};
  }
  if (result.after && options.expectedOrder) {
    result.actualOrder = result.after.sheets.map(s => s.id);
    result.orderMatches = result.actualOrder.join("\0") === options.expectedOrder.join("\0");
  }
  return result;
}
