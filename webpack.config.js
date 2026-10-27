/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const path = require("node:path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { loadPublicRuntimeConfig } = require("./runtime-config");

module.exports = (_environment, options = {}) => {
  const development = options.mode === "development";
  const runtimeConfig = loadPublicRuntimeConfig({
    defaultConfigFile: path.resolve(__dirname, "config.example.json"),
  });

  const config = {
    mode: development ? "development" : "production",
    devtool: development ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
    },
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },
    plugins: [
      new HtmlWebpackPlugin({
        chunks: ["taskpane"],
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
      }),
      new HtmlWebpackPlugin({
        chunks: [],
        filename: "commands.html",
        inject: false,
        template: "./src/commands/commands.html",
      }),
      new HtmlWebpackPlugin({
        chunks: [],
        filename: "support.html",
        inject: false,
        template: "./src/support.html",
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "./assets", to: "assets" },
          { from: "./src/taskpane/taskpane.css", to: "taskpane.css" },
          {
            from: "./config.example.json",
            to: "config.json",
            transform: () => `${JSON.stringify(runtimeConfig, null, 2)}\n`,
          },
          { from: "./config.schema.json", to: "config.schema.json" },
        ],
      }),
    ],
  };

  if (development) {
    config.watchOptions = {
      aggregateTimeout: 200,
      ignored: /node_modules/,
      poll: 1000,
    };
  }

  return config;
};
