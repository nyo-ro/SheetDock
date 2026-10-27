const path = require("node:path");
const fs = require("node:fs");
const os = require("node:os");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
// Based on the OfficeDev/Office-Addin-TaskPane-React configuration.
module.exports = (_env, argv) => {
  const config = {
    entry: { taskpane: "./src/taskpane/index.tsx" },
    output: { path: path.resolve(__dirname, "dist"), filename: "[name].[contenthash].js", clean: true },
    devtool: argv.mode === "development" ? "source-map" : false,
    resolve: { extensions: [".tsx", ".ts", ".js"] },
    module: { rules: [
      { test: /\.tsx?$/, exclude: /node_modules/, use: "ts-loader" },
      { test: /\.css$/, use: ["style-loader", "css-loader"] }
    ] },
    plugins: [
      new HtmlWebpackPlugin({ filename: "taskpane.html", template: "src/taskpane/taskpane.html" }),
      new CopyWebpackPlugin({ patterns: [
        { from: "assets", to: "assets" }, { from: "public", to: "." },
        { from: "manifest.xml", to: "manifest.xml" }
      ] })
    ],
    performance: { maxAssetSize: 600000, maxEntrypointSize: 600000 }
  };
  if (process.env.WEBPACK_SERVE) {
    // Never create or trust certificates implicitly.
    const certDir = process.env.SHEETDOCK_CERT_DIR || path.join(os.homedir(), ".office-addin-dev-certs");
    const key = path.join(certDir, "localhost.key");
    const cert = path.join(certDir, "localhost.crt");
    if (!fs.existsSync(key) || !fs.existsSync(cert)) {
      throw new Error("HTTPS証明書がありません。docs/development.md の証明書設定を確認してください。");
    }
    config.devServer = {
      host: "localhost", port: 3000, hot: true,
      server: { type: "https", options: { key: fs.readFileSync(key), cert: fs.readFileSync(cert) } },
      headers: { "Access-Control-Allow-Origin": "*" }
    };
  }
  return config;
};
