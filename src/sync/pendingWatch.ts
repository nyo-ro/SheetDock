import { diagnose } from "../diagnostics";
// 通知だけを期限切れにする。実APIのPromise/ロックは完了まで保持する。
export class PendingWatch {
  private entries=new Set<{expired:boolean;timer:ReturnType<typeof setTimeout>}>();
  private stopped=false;
  constructor(private changed:(stalled:boolean)=>void,private settled:()=>void,private timeout=15000){}
  async run<T>(label:string,work:(valid:()=>boolean)=>Promise<T>):Promise<T>{
    const entry={expired:false,timer:undefined as unknown as ReturnType<typeof setTimeout>};
    entry.timer=setTimeout(()=>{entry.expired=true;diagnose(label+".stalled");if(!this.stopped)this.changed(true);},this.timeout);
    this.entries.add(entry);
    try{return await work(()=>!this.stopped && !entry.expired);}
    finally{clearTimeout(entry.timer);this.entries.delete(entry);if(!this.stopped){this.changed([...this.entries].some(e=>e.expired));if(entry.expired){diagnose(label+".lateSettled");this.settled();}}}
  }
  stop():void{this.stopped=true;this.entries.forEach(e=>clearTimeout(e.timer));this.entries.clear();}
}
