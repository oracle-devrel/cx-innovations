/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const taskpaneHtml = fs.readFileSync(
  path.join(__dirname, "..", "src", "taskpane", "taskpane.html"),
  "utf8"
);
const taskpaneScript = fs.readFileSync(
  path.join(__dirname, "..", "src", "taskpane", "taskpane.js"),
  "utf8"
);

test("connection form hides the fixed REST path", () => {
  assert.doesNotMatch(taskpaneHtml, /id="rest-path"/);
  assert.doesNotMatch(taskpaneHtml, /REST base path/);
  assert.match(taskpaneHtml, /id="environment-url"/);
  assert.match(taskpaneHtml, /id="username"/);
  assert.match(taskpaneHtml, /id="password"/);
  assert.match(taskpaneHtml, /option value="basic"/);
  assert.match(taskpaneHtml, /option value="bearer"/);
  assert.doesNotMatch(taskpaneHtml, /option value="none"/);
  assert.doesNotMatch(taskpaneHtml, /remember-connection|Remember environment/);
  assert.doesNotMatch(taskpaneScript, /localStorage|CONNECTION_STORAGE_KEY/);
});
