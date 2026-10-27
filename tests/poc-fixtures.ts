import { vi } from "vitest";
import { Action, BookSnapshot } from "../src/poc/model";
export function book(): BookSnapshot {
  return {totalSheets:5,name:"SheetDock-PoC.xlsx",readOnly:false,structureProtected:false,activeId:"a",capturedAt:"test",
    sheets:["a","b","c","d","e"].map((id,position) => ({id,name:id,position,visibility:"Visible",protected:false,isActive:id === "a"}))};
}
export function memoryAdapter(initial = book()) {
  let current = structuredClone(initial);
  const adapter = {
    read: vi.fn(async () => structuredClone(current)),
    apply: vi.fn(async (action: Action) => {
      const index = current.sheets.findIndex(s => s.id === action.id);
      if (index < 0) throw new Error("ItemNotFound");
      switch(action.kind) {
        case "visibility": current.sheets[index].visibility = action.visibility; break;
        case "delete": current.sheets.splice(index,1); break;
        case "copy": current.sheets.splice(index+1,0,{...current.sheets[index],id:"new-id",name:"copy"}); break;
        case "move": { const [s] = current.sheets.splice(index,1); current.sheets.splice(action.position,0,s); break; }
      }
      current.sheets.forEach((s,position) => {s.position = position;});
    }),
    change: (mutate: (b: BookSnapshot) => void) => mutate(current)
  };
  return adapter;
}
