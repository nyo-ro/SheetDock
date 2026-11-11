const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert/strict");
const webpack = require("webpack");
const { JSDOM } = require("jsdom");
const root = path.resolve(__dirname, "..");
const config = require("../webpack.config.cjs")({ pages: true }, { mode: "production" });
const compiler = webpack({ ...config, mode: "production" });
compiler.run((error, stats) => {
  compiler.close(closeError => {
    try {
      if (error || closeError) throw error || closeError;
      const summary = stats.toJson({ all: false, errors: true, warnings: true, modules: true, nestedModules: true });
      fs.mkdirSync(path.join(root, ".local-tools"), { recursive: true });
      fs.writeFileSync(path.join(root, ".local-tools/pages-stats.json"), JSON.stringify(summary));
      if (stats.hasErrors()) throw new Error(JSON.stringify(summary.errors));
      function checkModules(modules) {
        for (const item of modules || []) {
          assert(!/[\\/]src[\\/](dev|poc)[\\/]/.test(item.name || ""), "Development module in Pages build");
          checkModules(item.modules);
        }
      }
      checkModules(summary.modules);
      const dist = path.join(root, "dist");
      const read = name => fs.readFileSync(path.join(dist, name), "utf8");
      const manifest = read("manifest.xml");
      assert.equal(manifest, fs.readFileSync(path.join(root, "manifest.pages.xml"), "utf8"));
      assert(!/localhost|127\.0\.0\.1/.test(manifest), "Local URL in Pages manifest");
      assert(manifest.includes('<Set Name="ExcelApi" MinVersion="1.17" />'));
      assert(fs.readFileSync(path.join(root, "manifest.xml"), "utf8").includes("https://localhost:3000"));
      const origin = "https://nyo-ro.github.io";
      const base = origin + "/SheetDock/";
      function checkAsset(reference, from = "taskpane.html") {
        const url = new URL(reference, base + from);
        if (url.origin !== origin) return;
        assert(url.pathname.startsWith("/SheetDock/"), "Asset outside Pages project path");
        const name = decodeURIComponent(url.pathname.slice("/SheetDock/".length)) || "index.html";
        assert(fs.statSync(path.join(dist, name)).isFile(), "Missing Pages asset: " + name);
      }
      for (const url of manifest.match(/https:\/\/nyo-ro\.github\.io[^"<>\s]*/g) || []) {
        if (url !== origin) checkAsset(url);
      }
      for (const file of ["index.html", "help.html", "taskpane.html", "commands.html"]) {
        const document = new JSDOM(read(file)).window.document;
        for (const element of document.querySelectorAll("[src], [href]")) {
          checkAsset(element.getAttribute("src") ?? element.getAttribute("href"), file);
        }
        if (file === "index.html") assert(document.querySelector('a[href="manifest.xml"][download]'));
      }
      for (const name of fs.readdirSync(dist).filter(name => name.endsWith(".js"))) {
        for (const marker of ["sheetDockDiagnostics", "isolated-read-probe", "sheetdock.readProbe", "manual.editStart"]) {
          assert(!read(name).includes(marker), "Development marker in Pages JavaScript");
        }
      }
      fs.writeFileSync(path.join(dist, ".nojekyll"), "");
      console.log(stats.toString({ colors: false, modules: false }));
      console.log("Pages checks passed: manifest, /SheetDock/ assets, download link, development-code exclusion.");
    } catch (failure) { console.error(failure); process.exitCode = 1; }
  });
});
