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
      {state.error && <div role="alert" className="error">{state.error}<br />表示内容を確認できません。手動更新、またはタスクペインを開き直してください。</div>}
      <p role="status">{state.hasSnapshot ? `${visible.length} / ${state.sheets.length} シート` : "一覧は未取得です。"}</p>
      <ul className="sheets" aria-label="ワークシート一覧">
        {visible.map(sheet => <li key={sheet.id}>
          <button type="button" className="sheet" aria-current={sheet.isActive ? "true" : undefined}
            disabled={state.busy || sheet.visibility !== "Visible"} onClick={() => void state.activate(sheet.id)}>
            <span className="sheet-name">{sheet.name}</span>
            <span className="sheet-state">{sheet.isActive ? "● アクティブ" : sheet.visibility === "Hidden" ? "非表示" : sheet.visibility === "VeryHidden" ? "完全非表示" : ""}</span>
          </button>
        </li>)}
      </ul>
      {state.hasSnapshot && !visible.length && <p>該当するワークシートはありません。</p>}
      <footer>Excel側の変更は「手動更新」で反映します。非表示シートは切り替えできません。独立したグラフシートは対象外です。</footer>
    </main>
  </FluentProvider>;
}
