/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const assert = require("node:assert/strict");
const test = require("node:test");
const {
  assertNoSecretConfigKeys,
  createPublicRuntimeConfig,
  parseBoolean,
} = require("../runtime-config");

function baseConfig() {
  return {
    api: {
      environmentUrl: "",
      restPath: "/crmRestApi/resources/11.13.18.05",
      revenueItemLimit: 20,
    },
    auth: {
      allowModeSelection: true,
      mode: "basic",
    },
    workbook: {
      maxRows: 20,
    },
  };
}

test("applies only allowlisted public startup overrides", () => {
  const input = baseConfig();
  const config = createPublicRuntimeConfig(input, {
    CX_API_ENVIRONMENT_URL: "https://sales.example.test/",
    CX_API_REST_PATH: "/customRestApi/",
    CX_AUTH_ALLOW_MODE_SELECTION: "false",
    CX_AUTH_MODE: "bearer",
    CX_API_SEND_COOKIES: "true",
    CX_RESOURCES_OPPORTUNITIES: "/must-not-be-applied",
    UNRELATED_SETTING: "must-not-be-applied",
  });

  assert.equal(config.api.environmentUrl, "https://sales.example.test");
  assert.equal(config.api.restPath, "/customRestApi");
  assert.equal(config.api.sendCookies, undefined);
  assert.equal(config.auth.allowModeSelection, false);
  assert.equal(config.auth.mode, "bearer");
  assert.equal(config.resources.opportunities, "/opportunities");
  assert.equal(config.UNRELATED_SETTING, undefined);
  assert.equal(input.api.environmentUrl, "", "startup overrides must not mutate the base file");
});

test("accepts only explicit true and false boolean startup values", () => {
  assert.equal(parseBoolean("CX_FLAG", " TRUE "), true);
  assert.equal(parseBoolean("CX_FLAG", "false"), false);
  assert.throws(() => parseBoolean("CX_FLAG", "1"), /CX_FLAG must be true or false/);
  assert.throws(() => parseBoolean("CX_FLAG", "yes"), /CX_FLAG must be true or false/);
  assert.throws(
    () => createPublicRuntimeConfig(baseConfig(), { CX_AUTH_ALLOW_MODE_SELECTION: "enabled" }),
    /CX_AUTH_ALLOW_MODE_SELECTION must be true or false/
  );
});

test("applies overrides when a base-file section has the wrong type", () => {
  const config = createPublicRuntimeConfig(
    {
      api: "invalid",
      auth: [],
    },
    {
      CX_API_ENVIRONMENT_URL: "https://sales.example.test",
      CX_AUTH_MODE: "bearer",
    }
  );

  assert.equal(config.api.environmentUrl, "https://sales.example.test");
  assert.equal(config.auth.mode, "bearer");
});

test("rejects removed or unknown authentication modes", () => {
  assert.throws(
    () => createPublicRuntimeConfig(baseConfig(), { CX_AUTH_MODE: "none" }),
    /Unsupported authentication mode: none/
  );
});

test("rejects a non-object base configuration", () => {
  assert.throws(
    () => createPublicRuntimeConfig(["invalid"], {}),
    /configuration must be a JSON object/
  );
});

test("rejects secret environment variables without echoing their values", () => {
  const secret = "do-not-repeat-this-password";

  assert.throws(
    () => createPublicRuntimeConfig(baseConfig(), { CX_PASSWORD: secret }),
    (error) => {
      assert.match(error.message, /CX_PASSWORD/);
      assert.doesNotMatch(error.message, new RegExp(secret));
      return true;
    }
  );
});

test("rejects secret-looking keys anywhere in public configuration", () => {
  const secret = "do-not-repeat-this-token";
  const config = baseConfig();
  config.sync = {
    extraRevenueFields: {
      bearer_token: secret,
    },
  };

  assert.throws(
    () => assertNoSecretConfigKeys(config),
    (error) => {
      assert.match(error.message, /sync\.extraRevenueFields\.bearer_token/);
      assert.doesNotMatch(error.message, new RegExp(secret));
      return true;
    }
  );
});
