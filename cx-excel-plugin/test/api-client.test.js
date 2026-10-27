/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

const test = require("node:test");
const assert = require("node:assert/strict");
const { ApiClient, ApiError, extractResourceKey } = require("../src/taskpane/api-client");
const { normalizeConfig } = require("../src/taskpane/config");

function config(overrides = {}) {
  return normalizeConfig({
    api: {
      environmentUrl: "https://sales.example.test",
      pageSize: 2,
      revenueItemLimit: 3,
      ...overrides,
    },
    workbook: { maxRows: 3 },
  });
}

function bearerAuth() {
  return { mode: "bearer", token: "test-access-token" };
}

test("sends Basic authentication without persisting credentials in configuration", async () => {
  let captured;
  const fetchImplementation = async (url, options) => {
    captured = { options, url };
    return new Response(JSON.stringify({ items: [], hasMore: false }), {
      headers: { "Content-Type": "application/json" },
      status: 200,
    });
  };
  const client = new ApiClient(
    config(),
    { mode: "basic", username: "üser", password: "päss" },
    fetchImplementation
  );

  await client.testConnection();

  assert.match(captured.options.headers.Authorization, /^Basic /);
  assert.equal(captured.options.credentials, "omit");
  assert.match(captured.url, /opportunities/);
  assert.equal(client.auth, undefined);
});

test("sends Bearer authentication with the provided access token", async () => {
  let capturedOptions;
  const client = new ApiClient(config(), bearerAuth(), async (_url, options) => {
    capturedOptions = options;
    return new Response(JSON.stringify({ items: [], hasMore: false }), { status: 200 });
  });

  await client.testConnection();

  assert.equal(capturedOptions.headers.Authorization, "Bearer test-access-token");
  assert.equal(capturedOptions.credentials, "omit");
});

test("rejects authentication modes other than Basic or Bearer", () => {
  assert.throws(
    () => new ApiClient(config(), { mode: "none" }, async () => new Response()),
    /Unsupported authentication mode: none/
  );
});

test("rejects a missing or whitespace-only Bearer access token", () => {
  assert.throws(
    () => new ApiClient(config(), { mode: "bearer", token: "   " }, async () => new Response()),
    /Enter an access token/
  );
});

test("identifies a likely CORS failure when a credential-free probe reaches the service", async () => {
  const calls = [];
  const client = new ApiClient(config(), bearerAuth(), async (url, options) => {
    calls.push({ options, url });
    if (calls.length === 1) {
      throw new TypeError("Load failed");
    }
    return new Response(null, { status: 401 });
  });

  await assert.rejects(() => client.testConnection(), /service appears reachable.*CORS/i);

  assert.equal(calls.length, 2);
  assert.equal(calls[1].url, calls[0].url);
  assert.equal(calls[1].options.credentials, "omit");
  assert.equal(calls[1].options.method, "GET");
  assert.equal(calls[1].options.mode, "no-cors");
  assert.equal(calls[1].options.headers, undefined);
  assert.equal(calls[1].options.body, undefined);
});

test("reports unconfirmed reachability when both the REST request and probe fail", async () => {
  const client = new ApiClient(config(), bearerAuth(), async () => {
    throw new TypeError("Load failed");
  });

  await assert.rejects(() => client.testConnection(), /Could not verify network access/);
});

test("does not probe again after a request timeout", async () => {
  let callCount = 0;
  const client = new ApiClient(config(), bearerAuth(), async () => {
    callCount += 1;
    const error = new Error("Aborted");
    error.name = "AbortError";
    throw error;
  });

  await assert.rejects(() => client.testConnection(), /Request timed out/);
  assert.equal(callCount, 1);
});

test("paginates Oracle collection responses up to the configured limit", async () => {
  const offsets = [];
  const pages = [
    { hasMore: true, items: [{ OptyNumber: "1" }, { OptyNumber: "2" }] },
    { hasMore: false, items: [{ OptyNumber: "3" }] },
  ];
  const client = new ApiClient(config({ opportunityLimit: 3 }), bearerAuth(), async (url) => {
    offsets.push(new URL(url).searchParams.get("offset"));
    return new Response(JSON.stringify(pages.shift()), { status: 200 });
  });

  const opportunities = await client.getOpportunities();

  assert.deepEqual(
    opportunities.map((item) => item.OptyNumber),
    ["1", "2", "3"]
  );
  assert.deepEqual(offsets, ["0", "2"]);
});

test("maps configurable product-group fields to the workbook contract", async () => {
  let requestedUrl;
  const customConfig = normalizeConfig({
    api: { environmentUrl: "https://sales.example.test" },
    resources: {
      productGroupIdField: "HierProdGroupId",
      productGroupNameField: "HierProdGroupName",
      productGroups: "/productGroups",
    },
  });
  const client = new ApiClient(customConfig, bearerAuth(), async (url) => {
    requestedUrl = new URL(url);
    return new Response(
      JSON.stringify({
        hasMore: false,
        items: [{ HierProdGroupId: 42, HierProdGroupName: "Services" }],
      }),
      { status: 200 }
    );
  });

  const groups = await client.getProductGroups();

  assert.equal(groups[0].ProdGroupId, 42);
  assert.equal(groups[0].ProdGroupName, "Services");
  assert.equal(requestedUrl.searchParams.get("fields"), "HierProdGroupId,HierProdGroupName");
});

test("encodes opportunity identifiers in revenue-item paths", () => {
  const client = new ApiClient(config(), bearerAuth(), async () => new Response());
  assert.equal(client.revenueItemsPath("OP/1"), "/opportunities/OP%2F1/child/ChildRevenue");
});

test("uses the server-provided revenue resource key from self links", async () => {
  let requestedUrl;
  const client = new ApiClient(config(), bearerAuth(), async (url) => {
    requestedUrl = new URL(url);
    return new Response(
      JSON.stringify({
        hasMore: false,
        items: [
          {
            RevnId: 10,
            links: [
              {
                href: "https://sales.example.test/crmRestApi/resources/11.13.18.05/opportunities/OP-1/child/ChildRevenue/hash%2Fkey",
                rel: "self",
              },
            ],
          },
        ],
      }),
      { status: 200 }
    );
  });

  const items = await client.getRevenueItems("OP-1");

  assert.equal(items[0]._resourceKey, "hash/key");
  assert.equal(requestedUrl.searchParams.get("onlyData"), "false");
});

test("rejects revenue records without a navigable resource key", () => {
  assert.throws(() => extractResourceKey({ RevnId: 10 }), /resource link/);
});

test("resolves relative revenue resource links against the configured environment", () => {
  const key = extractResourceKey(
    {
      links: [{ href: "/crmRestApi/resources/revenue/relative%2Fkey", rel: "self" }],
    },
    "https://sales.example.test"
  );

  assert.equal(key, "relative/key");
});

test("rejects incomplete successful batch responses", async () => {
  const client = new ApiClient(
    config(),
    bearerAuth(),
    async () => new Response(JSON.stringify({ parts: [] }), { status: 200 })
  );

  await assert.rejects(
    () => client.syncRevenueItems([{ id: "part1" }]),
    /incomplete or invalid batch response/
  );
});

test("turns structured Oracle error details into ApiError", async () => {
  const client = new ApiClient(
    config(),
    bearerAuth(),
    async () =>
      new Response(JSON.stringify({ "o:errorDetails": [{ detail: "Insufficient privilege" }] }), {
        status: 403,
      })
  );

  await assert.rejects(
    () => client.testConnection(),
    (error) =>
      error instanceof ApiError &&
      error.status === 403 &&
      error.message === "Insufficient privilege"
  );
});
