import { afterEach, expect, it, vi } from "vitest";
import { diagnose, diagnosticSnapshot, measured } from "../src/diagnostics";
afterEach(()=>vi.unstubAllEnvs());
it("診断は開発時のみ記録し機密を含むエラー本文を除外",async()=>{
 vi.stubEnv("NODE_ENV","development");const error=Object.assign(new Error("secret-sheet-name"),{code:"InvalidOperation",debugInfo:{errorLocation:"Worksheet.name",fullStatements:["secret-book"]}});
 await expect(measured("test.read",async()=>{throw error;})).rejects.toBe(error);const snapshot=diagnosticSnapshot();expect(snapshot.counts["test.read.start"]).toBe(1);expect(snapshot.counts["test.read.error"]).toBe(1);expect(snapshot.counts["test.read.end"]).toBe(1);expect(JSON.stringify(snapshot)).not.toContain("secret");expect(snapshot.entries.some(row=>row.code==="InvalidOperation")).toBe(true);
});
it("本番では診断行を追加しない",()=>{vi.stubEnv("NODE_ENV","production");const before=diagnosticSnapshot();diagnose("production");expect(diagnosticSnapshot()).toEqual(before);});
