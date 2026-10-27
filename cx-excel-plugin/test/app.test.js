/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const { once } = require("node:events");
const { createApp, startServer } = require("../app");

async function listen(app, context) {
  const server = app.listen(0, "127.0.0.1");
  context.after(() => new Promise((resolve) => server.close(resolve)));
  await new Promise((resolve, reject) => {
    server.once("listening", resolve);
    server.once("error", reject);
  });
  return `http://127.0.0.1:${server.address().port}`;
}

test("serves a hardened health endpoint", async (context) => {
  const baseUrl = await listen(createApp(), context);
  const response = await fetch(`${baseUrl}/healthz`);

  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), { status: "ok" });
  assert.equal(response.headers.get("x-content-type-options"), "nosniff");
  assert.equal(response.headers.get("referrer-policy"), "no-referrer");
  assert.equal(response.headers.get("x-powered-by"), null);
});

test("serves startup configuration dynamically without caching it", async (context) => {
  const runtimeConfig = {
    api: {
      environmentUrl: "https://sales.example.test",
      restPath: "/crmRestApi/resources/11.13.18.05",
    },
    auth: {
      allowModeSelection: false,
      mode: "bearer",
    },
  };
  const baseUrl = await listen(
    createApp({
      runtimeConfig,
      staticDirectory: path.join(__dirname, "missing-static-directory"),
    }),
    context
  );

  const response = await fetch(`${baseUrl}/config.json`);

  assert.equal(response.status, 200);
  assert.equal(response.headers.get("cache-control"), "no-store");
  assert.match(response.headers.get("content-type"), /^application\/json/);
  assert.deepEqual(await response.json(), runtimeConfig);
});

test("loads a public config file and environment overrides when the server starts", async (context) => {
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), "cx-add-in-config-"));
  const configFile = path.join(temporaryDirectory, "public-config.json");
  fs.writeFileSync(
    configFile,
    JSON.stringify({
      api: {
        environmentUrl: "https://file.example.test",
        restPath: "/fileRestApi",
        revenueItemLimit: 10,
      },
      auth: { mode: "basic" },
      workbook: { maxRows: 10 },
    })
  );

  const server = startServer({
    environment: {
      CX_API_ENVIRONMENT_URL: "https://startup.example.test",
      CX_AUTH_MODE: "bearer",
      CX_CONFIG_FILE: configFile,
      PORT: "0",
    },
    log() {},
    staticDirectory: path.join(temporaryDirectory, "static"),
  });
  context.after(
    () =>
      new Promise((resolve) => {
        server.close(() => {
          fs.rmSync(temporaryDirectory, { force: true, recursive: true });
          resolve();
        });
      })
  );
  await once(server, "listening");

  const response = await fetch(`http://127.0.0.1:${server.address().port}/config.json`);
  const config = await response.json();

  assert.equal(response.status, 200);
  assert.equal(config.api.environmentUrl, "https://startup.example.test");
  assert.equal(config.api.restPath, "/fileRestApi");
  assert.equal(config.auth.mode, "bearer");
});
