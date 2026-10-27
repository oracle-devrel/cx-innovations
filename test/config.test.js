/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const test = require("node:test");
const assert = require("node:assert/strict");
const {
  normalizeConfig,
  normalizeEnvironmentUrl,
  normalizeRestPath,
} = require("../src/taskpane/config");

test("normalizes safe service URLs and REST paths", () => {
  assert.equal(normalizeEnvironmentUrl("https://example.test/"), "https://example.test");
  assert.equal(normalizeEnvironmentUrl("http://localhost:3000"), "http://localhost:3000");
  assert.equal(
    normalizeRestPath("/crmRestApi/resources/11.13.18.05/"),
    "/crmRestApi/resources/11.13.18.05"
  );
  assert.throws(() => normalizeRestPath("crmRestApi/resources/11.13.18.05"), /absolute path/);
});

test("rejects unsafe or ambiguous service URLs", () => {
  assert.throws(() => normalizeEnvironmentUrl("http://example.test"), /HTTPS/);
  assert.throws(() => normalizeEnvironmentUrl("ftp://localhost"), /HTTPS/);
  assert.throws(() => normalizeEnvironmentUrl("https://user:secret@example.test"), /credentials/);
  assert.throws(() => normalizeEnvironmentUrl("https://example.test/api"), /only the origin/);
});

test("normalizes a complete runtime configuration", () => {
  const config = normalizeConfig({
    api: {
      environmentUrl: "https://sales.example.test",
      restPath: "/api/",
      revenueItemLimit: 20,
    },
    auth: { mode: "bearer" },
    resources: {
      productGroupIdField: "HierProdGroupId",
      productGroupNameField: "HierProdGroupName",
      productGroups: "/productGroups",
    },
    workbook: { maxRows: 20 },
    sync: { extraRevenueFields: { Example_c: "value" } },
    ui: {
      opportunityUrlTemplate: "https://sales.example.test/opportunity/{opportunityId}",
    },
  });

  assert.equal(config.api.restPath, "/api");
  assert.equal(config.auth.mode, "bearer");
  assert.equal(config.resources.productGroupIdField, "HierProdGroupId");
  assert.equal(config.resources.productGroupNameField, "HierProdGroupName");
  assert.equal(config.sync.extraRevenueFields.Example_c, "value");
});

test("supports only Basic and Bearer authentication modes", () => {
  assert.equal(normalizeConfig({ auth: { mode: "basic" } }).auth.mode, "basic");
  assert.equal(normalizeConfig({ auth: { mode: "bearer" } }).auth.mode, "bearer");
  assert.throws(
    () => normalizeConfig({ auth: { mode: "none" } }),
    /Unsupported authentication mode: none/
  );
});

test("rejects unsafe resource field names", () => {
  assert.throws(
    () => normalizeConfig({ resources: { productGroupIdField: "Id,SecretField" } }),
    /Invalid REST field name/
  );
});

test("rejects runtime configurations that cannot fit in the managed worksheet", () => {
  assert.throws(
    () => normalizeConfig({ api: { revenueItemLimit: 11 }, workbook: { maxRows: 10 } }),
    /cannot exceed/
  );
});

test("requires a placeholder in opportunity URL templates", () => {
  assert.throws(
    () => normalizeConfig({ ui: { opportunityUrlTemplate: "https://example.test/opportunity" } }),
    /must include/
  );
});

test("rejects protocol-relative and credential-bearing UI URLs", () => {
  assert.throws(
    () => normalizeConfig({ ui: { logoPath: "//untrusted.example/logo.png" } }),
    /relative path or an HTTPS URL/
  );
  assert.throws(
    () => normalizeConfig({ ui: { logoPath: "https://user:secret@example.test/logo.png" } }),
    /must not contain credentials/
  );
});
