import { ProbeKind, ProbeResult, ProbeTrace, SheetChoice } from "./api";

type Phase = "idle" | "preparing" | "ready" | "scheduled" | "running" | "succeeded" | "failed" | "stalled" | "settledLate" | "cancelled" | "disposed";
type Scope = "preparation" | "experiment";
export type ManualMark = "editStart" | "escape";
export interface ProbeView {
  phase: Phase; kind?: ProbeKind; choices: SheetChoice[]; count?: number;
  pending: boolean; waitingAt?: string; dueAt?: number; errorCode?: string;
}
export interface ProbeEntry {
  seq: number; time: number; isoTime: string; elapsedMs: number; stage: string;
  source: "manual" | "system"; durationMs?: number; code?: string; location?: string;
}
export interface ProbeDependencies {
  ready(): Promise<void>;
  prepare(trace: ProbeTrace): Promise<SheetChoice[]>;
  run(kind: ProbeKind, id: string | undefined, trace: ProbeTrace): Promise<ProbeResult>;
}
// このインスタンスは1ページにつき1つ。未応答Promiseを放棄して次のAPIを送らない。
export class ReadProbeController {
  readonly session = Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 10);
  private origin = performance.now();
  private state: ProbeView = { phase: "idle", choices: [], pending: false };
  private listeners = new Set<() => void>();
  private entries: ProbeEntry[] = [];
  private counts: Record<string, number> = {};
  private seq = 0;
  private scheduled?: ReturnType<typeof setTimeout>;
  private watchdog?: ReturnType<typeof setTimeout>;
  private disposed = false;
  constructor(private dependencies: ProbeDependencies, private timeoutMs = 15000) {
    this.record("runtime.init");
  }
  view(): ProbeView { return { ...this.state, choices: this.state.choices.map(s => ({ ...s })) }; }
  subscribe(listener: () => void): () => void { this.listeners.add(listener); return () => { this.listeners.delete(listener); }; }
  private changed() { if (!this.disposed) this.listeners.forEach(listener => listener()); }
  private record(stage: string, error?: unknown, durationMs?: number, source: "manual" | "system" = "system") {
    const value = error as { code?: unknown; debugInfo?: { errorLocation?: unknown } } | undefined;
    const code = typeof value?.code === "string" && /^[A-Za-z0-9_]{1,64}$/.test(value.code) ? value.code : undefined;
    const location = typeof value?.debugInfo?.errorLocation === "string" && /^(Workbook|Worksheet|WorksheetCollection|RequestContext|EventHandlers)(\.[A-Za-z]+)*$/.test(value.debugInfo.errorLocation) ? value.debugInfo.errorLocation : undefined;
    const time = Date.now();
    this.entries.push({ seq: ++this.seq, time, isoTime: new Date(time).toISOString(), elapsedMs: Math.round(performance.now() - this.origin), stage, source, durationMs, code, location });
    if (this.entries.length > 300) this.entries.shift();
    this.counts[stage] = (this.counts[stage] ?? 0) + 1;
    this.changed();
  }
  diagnostics() {
    const calls = (scope: Scope) => ({ run: this.counts[scope + ".api.run.start"] ?? 0, sync: this.counts[scope + ".api.sync.start"] ?? 0 });
    return { schema: 1, mode: "isolated-read-probe", session: this.session, kind: this.state.kind,
      status: this.state.phase, pending: this.state.pending, waitingAt: this.state.waitingAt,
      elapsedMs: Math.round(performance.now() - this.origin), apiCalls: { preparation: calls("preparation"), experiment: calls("experiment") },
      counts: { ...this.counts }, entries: this.entries.map(entry => ({ ...entry })) };
  }
  mark(mark: ManualMark) {
    if (!this.disposed && (mark === "editStart" || mark === "escape")) this.record("manual." + mark, undefined, undefined, "manual");
  }
  private trace(scope: Scope): ProbeTrace {
    return (stage, error, durationMs) => {
      if (stage === "api.run.start") this.state.waitingAt = scope + ".api.run";
      if (stage === "api.sync.start") this.state.waitingAt = scope + ".api.sync";
      if (stage === "api.sync.end") this.state.waitingAt = scope + ".api.run";
      if (stage === "api.run.end") this.state.waitingAt = undefined;
      this.record(scope + "." + stage, error, durationMs);
    };
  }
  async prepare(): Promise<void> {
    if (this.disposed || this.state.phase !== "idle") return;
    this.state.kind = "B";
    this.state.phase = "preparing";
    await this.execute("preparation", async trace => ({ count: 0, choices: await this.dependencies.prepare(trace) }));
  }
  schedule(kind: ProbeKind, knownId: string | undefined, delayMs: number): boolean {
    if (this.disposed || !["A", "B", "C"].includes(kind) || !Number.isInteger(delayMs) || delayMs < 0 || delayMs > 30000) return false;
    if (kind === "B") {
      if (this.state.phase !== "ready" || this.state.kind !== "B" || !this.state.choices.some(s => s.id === knownId)) return false;
    } else if (this.state.phase !== "idle") return false;
    this.state = { ...this.state, kind, phase: "scheduled", dueAt: Date.now() + delayMs };
    this.record("experiment.scheduled");
    this.scheduled = setTimeout(() => {
      this.scheduled = undefined;
      if (this.disposed) return;
      this.state.phase = "running"; this.state.dueAt = undefined;
      void this.execute("experiment", trace => this.dependencies.run(kind, knownId, trace));
    }, delayMs);
    return true;
  }
  cancelSchedule(): void {
    if (this.disposed || this.state.phase !== "scheduled") return;
    clearTimeout(this.scheduled); this.scheduled = undefined;
    this.state.phase = "cancelled"; this.state.dueAt = undefined;
    this.record("experiment.cancelledBeforeApi");
  }
  private async execute(scope: Scope, work: (trace: ProbeTrace) => Promise<ProbeResult>): Promise<void> {
    const started = performance.now(); let expired = false;
    this.state.pending = true; this.state.waitingAt = scope + ".ready";
    this.record(scope + ".start");
    this.watchdog = setTimeout(() => {
      expired = true;
      if (!this.disposed) { this.state.phase = "stalled"; this.record(scope + ".stalled"); }
    }, this.timeoutMs);
    try {
      this.record(scope + ".ready.start");
      await this.dependencies.ready();
      this.record(scope + ".ready.end");
      if (this.disposed || expired) { this.record(scope + ".cancelledBeforeApi"); return; }
      const result = await work(this.trace(scope));
      if (!this.disposed && !expired) {
        this.state.phase = scope === "preparation" ? "ready" : "succeeded";
        this.state.choices = result.choices ?? this.state.choices;
        if (scope === "experiment") this.state.count = result.count;
        this.record(scope + ".success");
      }
    } catch (error) {
      this.record(scope + ".error", error);
      const last = this.entries[this.entries.length - 1];
      if (!this.disposed) { this.state.errorCode = last.code ?? "UnknownError"; this.state.phase = "failed"; }
    } finally {
      clearTimeout(this.watchdog); this.watchdog = undefined;
      this.state.pending = false; this.state.waitingAt = undefined;
      if (!this.disposed && expired) { this.state.phase = "settledLate"; this.record(scope + ".lateSettled"); }
      this.record(scope + ".end", undefined, Math.round(performance.now() - started));
    }
  }
  dispose(): void {
    if (this.disposed) return;
    clearTimeout(this.scheduled); clearTimeout(this.watchdog);
    this.record("runtime.dispose"); this.disposed = true; this.state.phase = "disposed"; this.listeners.clear();
  }
}
