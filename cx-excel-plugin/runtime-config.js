/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const fs = require("node:fs");
const path = require("node:path");
const { normalizeConfig } = require("./src/taskpane/config");

const SECRET_ENVIRONMENT_VARIABLES = [
  "CX_ACCESS_TOKEN",
  "CX_BEARER_TOKEN",
  "CX_CLIENT_SECRET",
  "CX_PASSWORD",
  "CX_PRIVATE_KEY",
  "CX_REFRESH_TOKEN",
  "CX_USERNAME",
];

const SECRET_CONFIG_KEYS = new Set([
  "accesstoken",
  "authorization",
  "bearertoken",
  "clientsecret",
  "credentials",
  "password",
  "privatekey",
  "refreshtoken",
  "token",
  "username",
]);

function parseBoolean(name, value) {
  const normalized = String(value).trim().toLowerCase();
  if (normalized === "true") {
    return true;
  }
  if (normalized === "false") {
    return false;
  }
  throw new Error(`${name} must be true or false.`);
}

function assertNoSecretConfigKeys(value, trail = []) {
  if (!value || typeof value !== "object") {
    return;
  }
  for (const [key, nestedValue] of Object.entries(value)) {
    const normalizedKey = key.replaceAll(/[-_]/g, "").toLowerCase();
    if (SECRET_CONFIG_KEYS.has(normalizedKey)) {
      throw new Error(
        `Public runtime configuration must not contain secret field ${[...trail, key].join(".")}.`
      );
    }
    assertNoSecretConfigKeys(nestedValue, [...trail, key]);
  }
}

function setConfigValue(config, section, name, value) {
  if (!config[section] || typeof config[section] !== "object" || Array.isArray(config[section])) {
    config[section] = {};
  }
  config[section][name] = value;
}

function createPublicRuntimeConfig(baseConfig, environment = {}) {
  if (!baseConfig || typeof baseConfig !== "object" || Array.isArray(baseConfig)) {
    throw new Error("Public runtime configuration must be a JSON object.");
  }
  assertNoSecretConfigKeys(baseConfig);

  const suppliedSecretVariables = SECRET_ENVIRONMENT_VARIABLES.filter(
    (name) => environment[name] !== undefined && String(environment[name]).length > 0
  );
  if (suppliedSecretVariables.length) {
    throw new Error(
      `${suppliedSecretVariables.join(", ")} cannot be exposed through public runtime configuration. ` +
        "Enter direct-API credentials in the task pane or keep them server-side in an authenticated gateway."
    );
  }

  const config = JSON.parse(JSON.stringify(baseConfig));
  const stringOverrides = [
    ["CX_API_ENVIRONMENT_URL", "api", "environmentUrl"],
    ["CX_API_REST_PATH", "api", "restPath"],
    ["CX_AUTH_MODE", "auth", "mode"],
  ];
  for (const [environmentName, section, name] of stringOverrides) {
    if (environment[environmentName] !== undefined) {
      setConfigValue(config, section, name, String(environment[environmentName]));
    }
  }

  const booleanOverrides = [["CX_AUTH_ALLOW_MODE_SELECTION", "auth", "allowModeSelection"]];
  for (const [environmentName, section, name] of booleanOverrides) {
    if (environment[environmentName] !== undefined) {
      setConfigValue(
        config,
        section,
        name,
        parseBoolean(environmentName, environment[environmentName])
      );
    }
  }

  return normalizeConfig(config);
}

function loadPublicRuntimeConfig({
  defaultConfigFile = path.join(__dirname, "config.example.json"),
  environment = process.env,
} = {}) {
  const configuredPath = environment.CX_CONFIG_FILE || defaultConfigFile;
  const configFile = path.isAbsolute(configuredPath)
    ? configuredPath
    : path.resolve(process.cwd(), configuredPath);

  let parsed;
  try {
    parsed = JSON.parse(fs.readFileSync(configFile, "utf8"));
  } catch (error) {
    throw new Error(
      `Could not load public runtime configuration from ${configFile}: ${error.message}`
    );
  }
  return createPublicRuntimeConfig(parsed, environment);
}

module.exports = {
  SECRET_ENVIRONMENT_VARIABLES,
  assertNoSecretConfigKeys,
  createPublicRuntimeConfig,
  loadPublicRuntimeConfig,
  parseBoolean,
};
