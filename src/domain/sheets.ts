export type SheetVisibility = "Visible" | "Hidden" | "VeryHidden";
export interface SheetInfo {
  id: string;
  name: string;
  position: number;
  visibility: SheetVisibility;
  isActive: boolean;
}
export function filterSheets(sheets: readonly SheetInfo[], query: string): SheetInfo[] {
  const term = query.toLowerCase();
  return sheets.filter(sheet => sheet.name.toLowerCase().includes(term));
}
export function sortSheets(sheets: readonly SheetInfo[]): SheetInfo[] {
  return [...sheets].sort((a, b) => a.position - b.position);
}
