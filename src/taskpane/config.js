/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

/* global URL, fetch, module */

const AUTH_MODES = new Set(["basic", "bearer"]);
const LOOPBACK_HOSTS = new Set(["localhost", "127.0.0.1", "[::1]"]);

const DEFAULT_CONFIG = {
  api: {
    environmentUrl: "",
    restPath: "/crmRestApi/resources/11.13.18.05",
    frameworkVersion: "4",
    requestTimeoutMs: 30000,
    pageSize: 100,
    opportunityLimit: 25,
    productGroupLimit: 500,
    revenueItemLimit: 500,
  },
  auth: {
    mode: "basic",
    allowModeSelection: true,
  },
  resources: {
    opportunities: "/opportunities",
    productGroups: "/setupSalesCatalogs",
    productGroupIdField: "ProdGroupId",
    productGroupNameField: "ProdGroupName",
    revenueItems: "/opportunities/{opportunityNumber}/child/ChildRevenue",
  },
  workbook: {
    worksheetName: "CX Revenue Items",
    maxRows: 500,
    numberFormat: "#,##0.00",
  },
  sync: {
    extraRevenueFields: {},
  },
  ui: {
    title: "CX Sales",
    logoPath: "assets/oracle-o-icon-80.png",
    showDebug: false,
    opportunityUrlTemplate: "",
  },
};

function objectValue(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function integerValue(value, fallback, minimum, maximum) {
  if (value === undefined) {
    return fallback;
  }
  if (!Number.isInteger(value) || value < minimum || value > maximum) {
    throw new Error(`Expected an integer between ${minimum} and ${maximum}, received ${value}.`);
  }
  return value;
}

function booleanValue(value, fallback) {
  if (value === undefined) {
    return fallback;
  }
  if (typeof value !== "boolean") {
    throw new Error(`Expected a boolean, received ${typeof value}.`);
  }
  return value;
}

function stringValue(value, fallback) {
  if (value === undefined) {
    return fallback;
  }
  if (typeof value !== "string") {
    throw new Error(`Expected a string, received ${typeof value}.`);
  }
  return value.trim();
}

function normalizeEnvironmentUrl(value, allowEmpty = true) {
  const input = stringValue(value, "");
  if (!input && allowEmpty) {
    return "";
  }

  let url;
  try {
    url = new URL(input);
  } catch {
    throw new Error("Environment URL must be an absolute URL.");
  }

  const loopback = LOOPBACK_HOSTS.has(url.hostname);
  if (url.protocol !== "https:" && !(loopback && url.protocol === "http:")) {
    throw new Error(
      "Environment URL must use HTTPS (HTTP is allowed only for loopback development)."
    );
  }
  if (url.username || url.password || url.search || url.hash) {
    throw new Error("Environment URL must not contain credentials, a query string, or a fragment.");
  }
  if (url.pathname !== "/" && url.pathname !== "") {
    throw new Error("Environment URL must contain only the origin; put the API path in REST path.");
  }

  return url.origin;
}

function normalizeRestPath(value) {
  const input = stringValue(value, DEFAULT_CONFIG.api.restPath) || DEFAULT_CONFIG.api.restPath;
  if (
    !input.startsWith("/") ||
    input.startsWith("//") ||
    input.includes("?") ||
    input.includes("#")
  ) {
    throw new Error("REST path must be an absolute path without a query string or fragment.");
  }
  return input.replace(/\/+$/, "");
}

function normalizeResourcePath(value, fallback, requiredToken = "") {
  const input = stringValue(value, fallback) || fallback;
  if (
    !input.startsWith("/") ||
    input.startsWith("//") ||
    input.includes("://") ||
    input.includes("?") ||
    input.includes("#")
  ) {
    throw new Error(
      `Resource path must begin with a single slash and contain no URL or query: ${input}`
    );
  }
  if (requiredToken && !input.includes(requiredToken)) {
    throw new Error(`Resource path must include ${requiredToken}.`);
  }
  return input.replace(/\/+$/, "");
}

function normalizeFieldName(value, fallback) {
  const name = stringValue(value, fallback) || fallback;
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
    throw new Error(`Invalid REST field name: ${name}`);
  }
  return name;
}

function normalizeWorksheetName(value) {
  const name = stringValue(value, DEFAULT_CONFIG.workbook.worksheetName);
  const invalidCharacters = ["\\", "/", "*", "?", ":", "[", "]"];
  if (
    !name ||
    name.length > 31 ||
    invalidCharacters.some((character) => name.includes(character))
  ) {
    throw new Error("Worksheet name must be 1-31 characters and cannot contain \\ / * ? : [ or ].");
  }
  return name;
}

function normalizeExtraFields(value) {
  const fields = objectValue(value);
  for (const key of Object.keys(fields)) {
    if (["__proto__", "prototype", "constructor"].includes(key)) {
      throw new Error(`Unsupported extra revenue field name: ${key}`);
    }
  }
  return { ...fields };
}

function normalizeOpportunityUrlTemplate(value) {
  const template = stringValue(value, "");
  if (!template) {
    return "";
  }
  if (!template.includes("{opportunityId}") && !template.includes("{opportunityNumber}")) {
    throw new Error(
      "Opportunity URL template must include {opportunityId} or {opportunityNumber}."
    );
  }
  const sample = template
    .replaceAll("{opportunityId}", "1")
    .replaceAll("{opportunityNumber}", "OPTY-1");
  let url;
  try {
    url = new URL(sample);
  } catch {
    throw new Error("Opportunity URL template must be an absolute URL.");
  }
  if (url.protocol !== "https:" || url.username || url.password) {
    throw new Error("Opportunity URL template must use HTTPS.");
  }
  return template;
}

function normalizeLogoPath(value) {
  const logoPath = stringValue(value, DEFAULT_CONFIG.ui.logoPath);
  if (logoPath.startsWith("//") || /^(?:javascript|data):/i.test(logoPath)) {
    throw new Error("Logo path must be a relative path or an HTTPS URL.");
  }
  if (/^[a-z][a-z\d+.-]*:/i.test(logoPath) && !/^https:/i.test(logoPath)) {
    throw new Error("Logo path must be a relative path or an HTTPS URL.");
  }
  if (/^https:/i.test(logoPath)) {
    const url = new URL(logoPath);
    if (url.username || url.password) {
      throw new Error("Logo URL must not contain credentials.");
    }
  }
  return logoPath;
}

function normalizeConfig(value = {}) {
  const raw = objectValue(value);
  const api = objectValue(raw.api);
  const auth = objectValue(raw.auth);
  const resources = objectValue(raw.resources);
  const workbook = objectValue(raw.workbook);
  const sync = objectValue(raw.sync);
  const ui = objectValue(raw.ui);
  const authMode = stringValue(auth.mode, DEFAULT_CONFIG.auth.mode).toLowerCase();

  if (!AUTH_MODES.has(authMode)) {
    throw new Error(`Unsupported authentication mode: ${authMode}`);
  }

  const maximumRows = integerValue(workbook.maxRows, DEFAULT_CONFIG.workbook.maxRows, 1, 10000);
  const revenueItemLimit = integerValue(
    api.revenueItemLimit,
    DEFAULT_CONFIG.api.revenueItemLimit,
    1,
    10000
  );
  if (revenueItemLimit > maximumRows) {
    throw new Error("api.revenueItemLimit cannot exceed workbook.maxRows.");
  }

  return {
    api: {
      environmentUrl: normalizeEnvironmentUrl(api.environmentUrl),
      restPath: normalizeRestPath(api.restPath),
      frameworkVersion:
        stringValue(api.frameworkVersion, DEFAULT_CONFIG.api.frameworkVersion) ||
        DEFAULT_CONFIG.api.frameworkVersion,
      requestTimeoutMs: integerValue(
        api.requestTimeoutMs,
        DEFAULT_CONFIG.api.requestTimeoutMs,
        1000,
        120000
      ),
      pageSize: integerValue(api.pageSize, DEFAULT_CONFIG.api.pageSize, 1, 500),
      opportunityLimit: integerValue(
        api.opportunityLimit,
        DEFAULT_CONFIG.api.opportunityLimit,
        1,
        10000
      ),
      productGroupLimit: integerValue(
        api.productGroupLimit,
        DEFAULT_CONFIG.api.productGroupLimit,
        1,
        10000
      ),
      revenueItemLimit,
    },
    auth: {
      mode: authMode,
      allowModeSelection: booleanValue(
        auth.allowModeSelection,
        DEFAULT_CONFIG.auth.allowModeSelection
      ),
    },
    resources: {
      opportunities: normalizeResourcePath(
        resources.opportunities,
        DEFAULT_CONFIG.resources.opportunities
      ),
      productGroups: normalizeResourcePath(
        resources.productGroups,
        DEFAULT_CONFIG.resources.productGroups
      ),
      productGroupIdField: normalizeFieldName(
        resources.productGroupIdField,
        DEFAULT_CONFIG.resources.productGroupIdField
      ),
      productGroupNameField: normalizeFieldName(
        resources.productGroupNameField,
        DEFAULT_CONFIG.resources.productGroupNameField
      ),
      revenueItems: normalizeResourcePath(
        resources.revenueItems,
        DEFAULT_CONFIG.resources.revenueItems,
        "{opportunityNumber}"
      ),
    },
    workbook: {
      worksheetName: normalizeWorksheetName(workbook.worksheetName),
      maxRows: maximumRows,
      numberFormat:
        stringValue(workbook.numberFormat, DEFAULT_CONFIG.workbook.numberFormat) ||
        DEFAULT_CONFIG.workbook.numberFormat,
    },
    sync: {
      extraRevenueFields: normalizeExtraFields(sync.extraRevenueFields),
    },
    ui: {
      title: stringValue(ui.title, DEFAULT_CONFIG.ui.title) || DEFAULT_CONFIG.ui.title,
      logoPath: normalizeLogoPath(ui.logoPath),
      showDebug: booleanValue(ui.showDebug, DEFAULT_CONFIG.ui.showDebug),
      opportunityUrlTemplate: normalizeOpportunityUrlTemplate(ui.opportunityUrlTemplate),
    },
  };
}

async function loadRuntimeConfig(fetchImplementation = fetch, url = "config.json") {
  const response = await fetchImplementation(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Could not load runtime configuration (${response.status}).`);
  }
  return normalizeConfig(await response.json());
}

module.exports = {
  AUTH_MODES,
  DEFAULT_CONFIG,
  loadRuntimeConfig,
  normalizeConfig,
  normalizeEnvironmentUrl,
  normalizeFieldName,
  normalizeRestPath,
};
