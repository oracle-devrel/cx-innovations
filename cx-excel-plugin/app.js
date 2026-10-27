/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const https = require("node:https");
const path = require("node:path");
const express = require("express");
const { loadPublicRuntimeConfig } = require("./runtime-config");

function createApp({ runtimeConfig, staticDirectory = path.join(__dirname, "dist") } = {}) {
  const app = express();

  app.disable("x-powered-by");
  app.use((_request, response, next) => {
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");
    next();
  });
  app.get("/healthz", (_request, response) => response.status(200).json({ status: "ok" }));
  if (runtimeConfig) {
    app.get("/config.json", (_request, response) => {
      response.setHeader("Cache-Control", "no-store");
      response.status(200).json(runtimeConfig);
    });
  }
  app.use(
    express.static(staticDirectory, {
      setHeaders(response, filePath) {
        if (filePath.endsWith(".html") || filePath.endsWith("config.json")) {
          response.setHeader("Cache-Control", "no-store");
        }
      },
    })
  );

  return app;
}

function startServer({
  environment = process.env,
  httpsOptions,
  log = console.log,
  staticDirectory = path.join(__dirname, "dist"),
} = {}) {
  const port =
    environment.PORT === undefined || environment.PORT === "" ? 8080 : Number(environment.PORT);
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    throw new Error("PORT must be an integer between 0 and 65535.");
  }
  const runtimeConfig = loadPublicRuntimeConfig({
    defaultConfigFile: path.join(staticDirectory, "config.json"),
    environment,
  });
  const application = createApp({ runtimeConfig, staticDirectory });
  const listener = httpsOptions ? https.createServer(httpsOptions, application) : application;
  const server = listener.listen(port, () => {
    const address = server.address();
    const listeningPort = typeof address === "object" && address ? address.port : port;
    log(`CX Sales Excel Add-in static host listening on port ${listeningPort}`);
  });
  return server;
}

if (require.main === module) {
  startServer();
}

module.exports = { createApp, startServer };
