// イベントのまとまりを一回の取得に集約。取得中の要求は次の一回へ持ち越す。
export class CoalescedRefresh {
  private timer?: ReturnType<typeof setTimeout>;
  private dirty = false;
  private running = false;
  private stopped = false;
  constructor(private readonly refresh: () => Promise<void>, private readonly onError: (error: unknown) => void, private readonly delay = 150) {}
  request(): void {
    if (this.stopped) return;
    this.dirty = true;
    if (!this.running && this.timer === undefined) this.timer = setTimeout(() => { this.timer = undefined; void this.flush(); }, this.delay);
  }
  private async flush(): Promise<void> {
    if (this.stopped || !this.dirty) return;
    this.dirty = false; this.running = true;
    try { await this.refresh(); } catch (error) { if (!this.stopped) this.onError(error); }
    finally { this.running = false; if (this.dirty && !this.stopped) this.request(); }
  }
  stop(): void { this.stopped = true; this.dirty = false; if (this.timer !== undefined) clearTimeout(this.timer); }
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
