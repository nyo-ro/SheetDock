import { WorkbookState } from "../src/domain/workbookState";
export function workbook():WorkbookState{return {bookKey:"book-1",name:"Book.xlsx",readOnly:false,structureProtected:false,activeId:"a",totalSheets:3,
  sheets:["a","b","c"].map((id,position)=>({id,name:`Sheet${position+1}`,position,visibility:"Visible",isActive:id==="a",protected:false}))};}
