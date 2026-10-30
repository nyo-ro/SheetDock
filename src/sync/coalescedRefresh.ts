import { diagnose, measured } from "../diagnostics";
// 初回・手動・イベント取得を同じキューで実行。処理中の要求は後続一回へ。
export class CoalescedRefresh {
  private timer?: ReturnType<typeof setTimeout>;
  private dirty=false;
  private flight?:Promise<void>;
  private stopped=false;
  constructor(private readonly refresh:()=>Promise<void>,private readonly onError:(error:unknown)=>void,private readonly delay=150,private readonly label="refresh"){}
  request():void {
    if(this.stopped)return;diagnose(this.label+".request");if(this.flight || this.timer!==undefined)diagnose(this.label+".coalesced");this.dirty=true;
    this.schedule();
  }
  private schedule():void {if(!this.flight && this.timer===undefined)this.timer=setTimeout(()=>{this.timer=undefined;void this.flush();},this.delay);}
  runNow():Promise<void>{
    if(this.stopped)return Promise.resolve();diagnose(this.label+".request");this.dirty=true;
    if(this.timer!==undefined){clearTimeout(this.timer);this.timer=undefined;}
    return this.flight ?? this.flush();
  }
  private flush():Promise<void>{
    if(this.stopped || !this.dirty)return Promise.resolve();
    this.dirty=false;
    this.flight=Promise.resolve().then(()=>{if(!this.stopped)return measured(this.label,this.refresh);})
      .catch(error=>{if(!this.stopped)this.onError(error);})
      .finally(()=>{this.flight=undefined;if(this.dirty && !this.stopped)this.schedule();});
    return this.flight;
  }
  stop():void {this.stopped=true;this.dirty=false;if(this.timer!==undefined)clearTimeout(this.timer);diagnose(this.label+".stop");}
}
// 開始/終了要求を直列化。登録解除失敗時は再登録を禁止し二重登録を避ける。
export class SubscriptionLifecycle {
  private chain: Promise<void> = Promise.resolve();
  private dispose?: () => Promise<void>;
  private uncertain = false;
  constructor(private readonly register: () => Promise<() => Promise<void>>) {}
  setEnabled(enabled: boolean): Promise<void> {
    const task = this.chain.then(async () => {
      if (this.uncertain) throw new Error("イベント登録状態が不明です。ペインを再読み込みしてください。");
      if (enabled && !this.dispose) {
        try { this.dispose = await this.register(); } catch (error) { this.uncertain = true; throw error; }
      } else if (!enabled && this.dispose) {
        try { await this.dispose(); this.dispose = undefined; } catch (error) { this.uncertain = true; throw error; }
      }
    });
    this.chain = task.catch(() => undefined); return task;
  }
}
