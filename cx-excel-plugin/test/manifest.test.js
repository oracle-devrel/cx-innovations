/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const assert = require("node:assert/strict");

const sourceManifest = fs.readFileSync(path.join(__dirname, "..", "manifest.xml"), "utf8");

test("references bundled icons with the declared pixel dimensions", () => {
  const matches = sourceManifest.matchAll(
    /DefaultValue="https:\/\/localhost:8443\/assets\/oracle-o-icon-(\d+)\.png"/g
  );
  const sizes = new Set();

  for (const match of matches) {
    const declaredSize = Number(match[1]);
    const icon = fs.readFileSync(
      path.join(__dirname, "..", "assets", `oracle-o-icon-${declaredSize}.png`)
    );
    assert.equal(icon.subarray(1, 4).toString("ascii"), "PNG");
    assert.equal(icon.readUInt32BE(16), declaredSize);
    assert.equal(icon.readUInt32BE(20), declaredSize);
    sizes.add(declaredSize);
  }

  assert.deepEqual(
    [...sizes].sort((left, right) => left - right),
    [16, 32, 64, 80]
  );
});

test("uses the task-pane icon sizes required by the Excel Add-ins menu", () => {
  assert.match(
    sourceManifest,
    /<IconUrl DefaultValue="https:\/\/localhost:8443\/assets\/oracle-o-icon-32\.png"\/>/
  );
  assert.match(
    sourceManifest,
    /<HighResolutionIconUrl DefaultValue="https:\/\/localhost:8443\/assets\/oracle-o-icon-64\.png"\/>/
  );
  assert.match(sourceManifest, /<Version>1\.0\.2\.0<\/Version>/);
  assert.match(sourceManifest, /id="TaskpaneButton\.Label" DefaultValue="CX Sales"/);
  assert.doesNotMatch(sourceManifest, /DefaultValue="Revenue Items"/);
});
