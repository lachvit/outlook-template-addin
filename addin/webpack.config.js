const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// The manifests in source control always use this address; production builds replace it.
const DEV_URL = "https://localhost:3000/";

function withTrailingSlash(url) {
  return url.endsWith("/") ? url : `${url}/`;
}

async function getDevHttpsOptions() {
  const devCerts = require("office-addin-dev-certs");
  const options = await devCerts.getHttpsServerOptions();
  return { ca: options.ca, key: options.key, cert: options.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  // ADDIN_BASE_URL is the single setting that decides every public URL of the add-in.
  if (!dev && !process.env.ADDIN_BASE_URL) {
    throw new Error("Set ADDIN_BASE_URL (for example https://<org>.github.io/<repo>/) for production builds.");
  }
  const baseUrl = withTrailingSlash(dev ? DEV_URL : process.env.ADDIN_BASE_URL);
  if (!baseUrl.startsWith("https://")) {
    throw new Error(`ADDIN_BASE_URL must start with https:// (got ${baseUrl}).`);
  }
  const templatesBaseUrl = withTrailingSlash(process.env.TEMPLATES_BASE_URL || `${baseUrl}templates/`);
  const buildId = process.env.BUILD_ID || (dev ? "dev" : String(Date.now()));

  const rewriteUrls = (content) =>
    dev ? content : content.toString().split(DEV_URL).join(baseUrl);

  return {
    devtool: dev ? "source-map" : false,
    entry: { taskpane: "./src/taskpane/index.tsx" },
    output: {
      path: path.resolve(__dirname, "dist"),
      publicPath: baseUrl,
      filename: "[name].[contenthash].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".tsx", ".js"] },
    module: {
      rules: [{ test: /\.tsx?$/, exclude: /node_modules/, use: "ts-loader" }],
    },
    plugins: [
      new webpack.DefinePlugin({
        __TEMPLATES_BASE_URL__: JSON.stringify(templatesBaseUrl),
        __BUILD_ID__: JSON.stringify(buildId),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "manifest.json", to: "manifest.json", transform: rewriteUrls },
          { from: "manifest.xml", to: "manifest.xml", transform: rewriteUrls },
        ],
      }),
    ],
    performance: { hints: false },
    devServer: {
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: env.WEBPACK_SERVE ? await getDevHttpsOptions() : {} },
      // Serves ../templates at /templates so the task pane works locally (run the publish script first).
      static: [{ directory: path.resolve(__dirname, "../templates"), publicPath: "/templates" }],
    },
  };
};
