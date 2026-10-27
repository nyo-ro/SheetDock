import { SheetList } from "./SheetList";
import { Button, FluentProvider, Input, Spinner, webLightTheme } from "@fluentui/react-components";
import { filterSheets } from "../domain/sheets";
import { useWorksheets } from "../state/useWorksheets";
import "./styles.css";

export function App() {
  const state = useWorksheets();
  const visible = filterSheets(state.sheets, state.query);
  return <FluentProvider theme={webLightTheme}>
    <main aria-busy={state.busy}>
      <header><h1>SheetDock</h1><Button onClick={() => void state.refresh()} disabled={!state.ready || state.busy}>手動更新</Button></header>
      <label htmlFor="sheet-search">シート名で検索</label>
      <div className="search"><Input id="sheet-search" value={state.query} onChange={(_, data) => state.setQuery(data.value)} placeholder="名前の一部を入力" />
        <Button onClick={() => state.setQuery("")} disabled={!state.query}>クリア</Button></div>
      {state.busy && <Spinner size="tiny" label="Excelから取得中…" />}
      {state.error && <div role="alert" className="error">{state.error}<br />表示内容をExcelと照合し、必要なら手動更新してください。</div>}
      {state.syncError && <p role="alert">{state.syncError} 手動更新で再接続を試せます。</p>}
      <p role="status">{state.hasSnapshot ? `${visible.length} / ${state.sheets.length} シート` : "一覧は未取得です。"}</p>
      <SheetList queryKey={state.query} sheets={visible} snapshot={state.snapshot} busy={state.busy} disabledReason={state.moveDisabledReason}
        activate={id=>void state.activate(id)} move={(baseline,id,boundary,epoch)=>void state.move(baseline,id,boundary,epoch)} dragEpoch={state.dragEpoch} />
      {state.hasSnapshot && !visible.length && <p>該当するワークシートはありません。</p>}
      {state.moveDisabledReason && <p className="move-hint">{state.moveDisabledReason}</p>}
      <footer>Excel側の変更を自動反映します。取得に失敗した場合は手動更新してください。非表示シートは切り替えできません。</footer>
    </main>
  </FluentProvider>;
}
