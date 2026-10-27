/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const { spawnSync } = require("node:child_process");
const devCerts = require("office-addin-dev-certs");
const { startServer } = require("../app");

function buildDevelopmentBundle() {
  const result = spawnSync(
    process.execPath,
    [require.resolve("webpack-cli/bin/cli.js"), "--mode", "development"],
    {
      env: process.env,
      stdio: "inherit",
    }
  );
  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`The development bundle failed with exit code ${result.status}.`);
  }
}

async function main() {
  buildDevelopmentBundle();
  const httpsOptions = await devCerts.getHttpsServerOptions();
  startServer({
    environment: {
      ...process.env,
      PORT: process.env.PORT || "8443",
    },
    httpsOptions,
  });
}

main().catch((error) => {
  console.error(error.stack || error.message || String(error));
  process.exitCode = 1;
});
