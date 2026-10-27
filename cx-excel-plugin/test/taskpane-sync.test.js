/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const test = require("node:test");
const assert = require("node:assert/strict");
const originalOffice = global.Office;
global.Office = { onReady() {} };
const { buildBatchParts } = require("../src/taskpane/taskpane");
if (originalOffice === undefined) {
  delete global.Office;
} else {
  global.Office = originalOffice;
}

test("builds create and update batch parts with product-group IDs", () => {
  const originals = new Map([
    [
      "10",
      {
        ProdGroupId: 101,
        ProdGroupName: "Air",
        ProductType: "Group",
        Quantity: 1,
        RevnAmount: 100,
        UnitPrice: 100,
      },
    ],
  ]);
  const rows = [
    ["10", "Group", "Air", 2, 100, 200, ""],
    ["", "Group", "Ocean", 1, 250, 250, ""],
  ];

  const result = buildBatchParts({
    extraFields: { Example_c: "configured" },
    originals,
    productGroupIdsByName: new Map([
      ["Air", 101],
      ["Ocean", 202],
    ]),
    revenueItemsPath: "/opportunities/OP-1/child/ChildRevenue",
    rows,
  });

  assert.equal(result.parts.length, 2);
  assert.equal(result.parts[0].operation, "update");
  assert.equal(result.parts[0].path.endsWith("/10"), true);
  assert.equal(result.parts[0].payload.ProdGroupId, 101);
  assert.equal(result.parts[0].payload.Example_c, "configured");
  assert.equal(result.parts[1].operation, "create");
  assert.equal(result.parts[1].payload.ProdGroupId, 202);
});

test("does not send unchanged rows", () => {
  const original = {
    ProdGroupId: 101,
    ProdGroupName: "Air",
    ProductType: "Group",
    Quantity: 1,
    RevnAmount: 100,
    UnitPrice: 100,
  };
  const result = buildBatchParts({
    originals: new Map([["10", original]]),
    productGroupIdsByName: new Map([["Air", 101]]),
    revenueItemsPath: "/opportunities/OP-1/child/ChildRevenue",
    rows: [["10", "Group", "Air", 1, 100, 100, ""]],
  });

  assert.equal(result.parts.length, 0);
});

test("rejects unsupported product rows with an actionable row number", () => {
  assert.throws(
    () =>
      buildBatchParts({
        originals: new Map(),
        productGroupIdsByName: new Map(),
        revenueItemsPath: "/opportunities/OP-1/child/ChildRevenue",
        rows: [["", "Product", "Widget", 1, 10, 10, ""]],
      }),
    /Row 2.*only product-group/
  );
});

test("sends nulls when an existing numeric value is cleared", () => {
  const original = {
    ProdGroupId: 101,
    ProdGroupName: "Air",
    ProductType: "Group",
    Quantity: 1,
    RevnAmount: 100,
    UnitPrice: 100,
  };
  const result = buildBatchParts({
    originals: new Map([["opaque/key", original]]),
    productGroupIdsByName: new Map([["Air", 101]]),
    revenueItemsPath: "/opportunities/OP-1/child/ChildRevenue",
    rows: [["opaque/key", "Group", "Air", "", "", "", ""]],
  });

  assert.equal(result.parts[0].path.endsWith("/opaque%2Fkey"), true);
  assert.equal(result.parts[0].payload.Quantity, null);
  assert.equal(result.parts[0].payload.UnitPrice, null);
  assert.equal(result.parts[0].payload.RevnAmount, null);
});

test("rejects an unknown edited product group instead of reusing the original ID", () => {
  assert.throws(
    () =>
      buildBatchParts({
        originals: new Map([
          [
            "key",
            {
              ProdGroupId: 101,
              ProdGroupName: "Air",
              ProductType: "Group",
              Quantity: 1,
              RevnAmount: 100,
              UnitPrice: 100,
            },
          ],
        ]),
        productGroupIdsByName: new Map([["Air", 101]]),
        revenueItemsPath: "/opportunities/OP-1/child/ChildRevenue",
        rows: [["key", "Group", "Typo", 1, 100, 100, ""]],
      }),
    /was not found/
  );
});
