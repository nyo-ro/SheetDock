// 開発専用。ブック名・シート名・ID・URL・引数・生エラー本文は記録しない。
type Entry = { seq:number; time:number; stage:string; durationMs?:number; code?:string; location?:string };
const rows:Entry[]=[];const counts:Record<string,number>={};let seq=0;let consoleOutput=false;let consoleSecond=0,consoleCount=0;
const session=Date.now().toString(36)+Math.random().toString(36).slice(2,7);
export function diagnose(stage:string,error?:unknown,durationMs?:number):void {
  if(process.env.NODE_ENV!=="development")return;
  const e=error as {code?:unknown;debugInfo?:{errorLocation?:unknown}}|undefined;
  const code=typeof e?.code==="string" && /^[A-Za-z0-9_]{1,64}$/.test(e.code) ? e.code : undefined;
  const location=typeof e?.debugInfo?.errorLocation==="string" && /^(Workbook|Worksheet|WorksheetCollection|Functions|RequestContext|EventHandlers)(\.[A-Za-z]+)*$/.test(e.debugInfo.errorLocation) ? e.debugInfo.errorLocation : undefined;
  counts[stage]=(counts[stage]??0)+1;
  const entry={seq:++seq,time:Date.now(),stage,durationMs,code,location};
  rows.push(entry);if(rows.length>500)rows.shift();if(consoleOutput){const second=Math.floor(Date.now()/1000);if(second!==consoleSecond){consoleSecond=second;consoleCount=0;}if(consoleCount++<20)console.info("SheetDock diagnostic",JSON.stringify(entry));}
}
export async function measured<T>(stage:string,work:()=>Promise<T>):Promise<T>{
  if(process.env.NODE_ENV!=="development")return work();
  const start=performance.now();diagnose(stage+".start");
  try{return await work();}catch(error){diagnose(stage+".error",error);throw error;}
  finally{diagnose(stage+".end",undefined,Math.round(performance.now()-start));}
}
export const diagnosticSnapshot=()=>({session,counts:{...counts},entries:rows.map(row=>({...row}))});
if(process.env.NODE_ENV==="development" && typeof window!=="undefined"){
  Object.assign(window,{sheetDockDiagnostics:{snapshot:diagnosticSnapshot,setConsole:(enabled:boolean)=>{consoleOutput=enabled;},clear:()=>{rows.length=0;for(const key of Object.keys(counts))delete counts[key];}}});
  window.addEventListener("pagehide",()=>diagnose("runtime.pagehide"));
  window.addEventListener("pageshow",()=>diagnose("runtime.pageshow"));
  window.addEventListener("error",event=>diagnose("runtime.error",event.error));
  window.addEventListener("unhandledrejection",event=>diagnose("runtime.unhandledrejection",event.reason));
}
