const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// The manifests in source control always use this address; production builds replace it.
const DEV_URL = "https://localhost:3000/";
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DEFAULT_LIBRARY = "Email Templates";

// Optional local settings file (see .env.example). Real environment variables take precedence.
const envFile = path.resolve(__dirname, ".env");
if (fs.existsSync(envFile)) {
  process.loadEnvFile(envFile);
}

function withTrailingSlash(url) {
  return url.endsWith("/") ? url : `${url}/`;
}

/** Reads the Entra ID and SharePoint settings; production builds fail fast when one is missing or malformed. */
function readSettings(dev) {
  const settings = {
    AUTH_CLIENT_ID: (process.env.AUTH_CLIENT_ID || "").trim(),
    AUTH_TENANT_ID: (process.env.AUTH_TENANT_ID || "").trim(),
    SHAREPOINT_SITE_URL: (process.env.SHAREPOINT_SITE_URL || "").trim().replace(/\/+$/, ""),
    SHAREPOINT_LIBRARY: (process.env.SHAREPOINT_LIBRARY || DEFAULT_LIBRARY).trim(),
    SHAREPOINT_FOLDER: (process.env.SHAREPOINT_FOLDER || "").trim().replace(/^\/+|\/+$/g, ""),
  };
  const problems = [];
  for (const name of ["AUTH_CLIENT_ID", "AUTH_TENANT_ID"]) {
    if (!GUID.test(settings[name])) {
      problems.push(`${name} must be a GUID from the Entra ID app registration (got "${settings[name]}").`);
    }
  }
  if (!/^https:\/\/[^/]+\.sharepoint\.com(\/|$)/i.test(settings.SHAREPOINT_SITE_URL)) {
    problems.push(
      `SHAREPOINT_SITE_URL must look like https://<tenant>.sharepoint.com/sites/<site> (got "${settings.SHAREPOINT_SITE_URL}").`,
    );
  }
  if (problems.length > 0) {
    if (!dev) {
      throw new Error(`Invalid add-in settings:\n- ${problems.join("\n- ")}`);
    }
    console.warn(`\nAdd-in settings are incomplete; the task pane will show a setup message:\n- ${problems.join("\n- ")}\n`);
  }
  return settings;
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
  const settings = readSettings(dev);

  const rewriteUrls = (content) =>
    dev ? content : content.toString().split(DEV_URL).join(baseUrl);

  return {
    devtool: dev ? "source-map" : false,
    entry: { taskpane: "./src/taskpane/index.tsx", auth: "./src/auth/auth.ts" },
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
        __AUTH_CLIENT_ID__: JSON.stringify(settings.AUTH_CLIENT_ID),
        __AUTH_TENANT_ID__: JSON.stringify(settings.AUTH_TENANT_ID),
        __SHAREPOINT_SITE_URL__: JSON.stringify(settings.SHAREPOINT_SITE_URL),
        __SHAREPOINT_LIBRARY__: JSON.stringify(settings.SHAREPOINT_LIBRARY),
        __SHAREPOINT_FOLDER__: JSON.stringify(settings.SHAREPOINT_FOLDER),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/auth/auth.html",
        chunks: ["auth"],
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
    },
  };
};
