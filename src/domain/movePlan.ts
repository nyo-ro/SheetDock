// Sheaside WorksheetMoveRules.GroupOrderの境界規則をTypeScriptで再表現。
export function groupOrder(full: readonly string[], displayed: readonly string[], selected: readonly string[], boundary: number): string[] {
  const unique = (xs: readonly string[]) => xs.every(Boolean) && new Set(xs).size === xs.length;
  if (![full, displayed, selected].every(unique) || !displayed.length || !selected.length ||
      !Number.isInteger(boundary) || boundary < 0 || boundary > displayed.length ||
      selected.some(id => !displayed.includes(id)) || displayed.some(id => !full.includes(id)) ||
      full.filter(id => displayed.includes(id)).join("\0") !== displayed.join("\0")) {
    throw new Error("移動対象または一覧が変更されています。");
  }
  const gap = boundary < displayed.length ? full.indexOf(displayed[boundary]) : full.indexOf(displayed[displayed.length - 1]) + 1;
  const insertion = full.slice(0, gap).filter(id => !selected.includes(id)).length;
  const remaining = full.filter(id => !selected.includes(id));
  return [...remaining.slice(0, insertion), ...full.filter(id => selected.includes(id)), ...remaining.slice(insertion)];
}
export interface PositionWrite { id: string; position: number; orderAfter: string[] }
export function positionPlan(full: readonly string[], selected: readonly string[], desired: readonly string[]): PositionWrite[] {
  if (new Set(full).size !== full.length || full.length !== desired.length || new Set(desired).size !== desired.length ||
      desired.some(id => !full.includes(id)) || !selected.length || new Set(selected).size !== selected.length || selected.some(id => !full.includes(id))) throw new Error("不正な移動計画です。");
  const moving = full.filter(id => selected.includes(id));
  if (full.filter(id => !selected.includes(id)).join() !== desired.filter(id => !selected.includes(id)).join() ||
      moving.join() !== desired.filter(id => selected.includes(id)).join()) throw new Error("相対順序を変更できません。");
  const start = desired.indexOf(moving[0]);
  if (desired.slice(start, start + moving.length).join() !== moving.join()) throw new Error("移動対象は連続した配置にしてください。");
  if (full.join() === desired.join()) return [];
  const order = [...full]; const writes: PositionWrite[] = [];
  const move = (id: string, position: number) => {
    const from = order.indexOf(id); if (from === position) return;
    order.splice(from, 1); order.splice(position, 0, id);
    writes.push({ id, position, orderAfter: [...order] });
  };
  if (moving.length === 1) move(moving[0], start);
  else {
    // 選択シートだけを末尾へ集めてから最終位置へ。途中順序も計画に記録する。
    moving.forEach(id => move(id, full.length - 1));
    moving.forEach((id, index) => move(id, start + index));
  }
  if (order.join() !== desired.join()) throw new Error("移動計画の整合性エラーです。");
  return writes;
}
