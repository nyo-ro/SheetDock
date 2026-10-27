// 複数の利用者（通常UI・PoC）で1つの登録を共有する。
export class SharedEvents<T> {
  private listeners=new Map<symbol,(event:T)=>void>();
  private dispose?:()=>Promise<void>;
  private chain:Promise<unknown>=Promise.resolve();
  private uncertain=false;
  constructor(private register:(emit:(event:T)=>void)=>Promise<()=>Promise<void>>) {}
  private serial<R>(work:()=>Promise<R>):Promise<R>{const result=this.chain.then(work);this.chain=result.catch(()=>undefined);return result;}
  private async connect(){
    if(this.uncertain)throw new Error("イベント登録状態が不明です。ペインを開き直してください。");
    if(!this.dispose){try{this.dispose=await this.register(e=>this.listeners.forEach(l=>l(e)));}catch(e){if((e as {uncertain?:boolean})?.uncertain)this.uncertain=true;throw e;}}
  }
  subscribe(listener:(event:T)=>void):Promise<()=>Promise<void>>{
    const token=Symbol();return this.serial(async()=>{
      this.listeners.set(token,listener);
      try{await this.connect();}catch(e){this.listeners.delete(token);throw e;}
      let released=false;
      return ()=>this.serial(async()=>{
        if(released)return;released=true;this.listeners.delete(token);
        if(!this.listeners.size && this.dispose){try{await this.dispose();this.dispose=undefined;}catch(e){this.uncertain=true;throw e;}}
      });
    });
  }
  rebind():Promise<void>{return this.serial(async()=>{
    if(this.uncertain)throw new Error("イベント登録状態が不明です。ペインを開き直してください。");
    if(this.dispose){try{await this.dispose();this.dispose=undefined;}catch(e){this.uncertain=true;throw e;}}
    if(this.listeners.size)await this.connect();
  });}
}
