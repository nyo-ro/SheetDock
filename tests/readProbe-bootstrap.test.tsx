// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { startDevelopment } from "../src/dev/bootstrap";
import { isReadProbeMode } from "../src/dev/readProbe/mode";
const fake = vi.hoisted(() => ({ render: vi.fn(), createRoot: vi.fn() }));
vi.mock("react-dom/client", () => ({ createRoot: fake.createRoot }));
vi.mock("../src/dev/readProbe/mode", () => ({ isReadProbeMode: vi.fn(), reloadInMode: vi.fn() }));
vi.mock("../src/dev/readProbe/Panel", () => ({ ReadProbePanel: function ReadProbePanel() { return null; } }));
vi.mock("../src/poc/Panel", () => ({ Panel: function Panel() { return null; } }));
beforeEach(() => { vi.clearAllMocks(); fake.createRoot.mockReturnValue({render:fake.render}); document.body.innerHTML='<div id="root"></div>'; });
afterEach(() => { document.body.innerHTML=""; });
it("比較モードは専用ルートだけ起動し通常UI・PoCをマウントしない", async () => {
  vi.mocked(isReadProbeMode).mockReturnValue(true); const normal=vi.fn(); const root=document.getElementById("root")!;
  await startDevelopment(root,normal); expect(normal).not.toHaveBeenCalled(); expect(fake.createRoot).toHaveBeenCalledExactlyOnceWith(root);
  expect(fake.render).toHaveBeenCalledOnce(); expect(fake.render.mock.calls[0][0].type.name).toBe("ReadProbePanel"); expect(document.body.children).toHaveLength(1);
});
it("通常の開発起動は通常UIと既存PoCを維持し、比較モード入口を追加", async () => {
  vi.mocked(isReadProbeMode).mockReturnValue(false); const normal=vi.fn(); await startDevelopment(document.getElementById("root")!,normal);
  expect(normal).toHaveBeenCalledOnce(); expect(fake.createRoot).toHaveBeenCalledTimes(2);
  expect(fake.render.mock.calls.map(args=>args[0].type.name)).toEqual(["ProbeLauncher","Panel"]);
  expect(document.body.children).toHaveLength(3);
});
