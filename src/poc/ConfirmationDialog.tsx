import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { Action, BookSnapshot } from "./model";
import { ExperimentOptions } from "./experiment";
export interface OperationPlan { baseline: BookSnapshot; actions: Action[]; options: ExperimentOptions }
interface Props { plan: OperationPlan; onCancel(): void; onExecute(): void }
export function ConfirmationDialog({ plan, onCancel, onExecute }: Props) {
  const sheetName = (id: string) => plan.baseline.sheets.find(sheet => sheet.id === id)?.name ?? id;
  const describe = (action: Action) => {
    const target = `${sheetName(action.id)}（ID: ${action.id}）`;
    switch (action.kind) {
      case "move": return `${target} を位置 ${action.position + 1} へ移動`;
      case "visibility": return `${target} の表示状態を ${action.visibility} へ変更`;
      case "copy": return `${target} の直後へ複製（複製先の表示状態は実行後に確認）`;
      case "delete": return `${target} を削除`;
    }
  };
  return <FluentProvider theme={webLightTheme}>
    <Dialog open onOpenChange={(_, data) => { if (!data.open) onCancel(); }}>
      <DialogSurface style={{ maxWidth: "calc(100vw - 24px)", overflowWrap: "anywhere" }}>
        <DialogBody>
          <DialogTitle>検証操作の実行確認</DialogTitle>
          <DialogContent>
            <p>対象ブック：<strong>{plan.baseline.name}</strong></p>
            <p>実行する操作（{plan.actions.length}ステップ）</p>
            <ol>{plan.actions.map((action, index) => <li key={index}>{describe(action)}</li>)}</ol>
            {!plan.actions.length && <p>順序は変わりません。状態の再検証のみ行います。</p>}
            {plan.options.expectedOrder && <p>想定する最終順序：{plan.options.expectedOrder.map(sheetName).join(" → ")}</p>}
            <p>使い捨ての検証ブックであることを確認してください。実行直前に状態を再取得し、計画と異なる場合は中断します。</p>
            <p>削除を元に戻せる保証はありません。VeryHiddenの復元も保証しません。途中失敗時は変更が一部残る場合があります。</p>
            {plan.options.injectFailure && <p>合成失敗：ステップ {plan.options.injectFailure.step} の{plan.options.injectFailure.when === "before" ? "書き込み前" : "書き込み後・確認前"}にエラーを注入します。</p>}
          </DialogContent>
          <DialogActions>
            <Button onClick={onCancel}>キャンセル</Button>
            <Button appearance="primary" onClick={onExecute}>実行</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  </FluentProvider>;
}
