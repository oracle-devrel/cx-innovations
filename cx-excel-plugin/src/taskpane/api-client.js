/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

/* global AbortController, URL, clearTimeout, fetch, globalThis, module, setTimeout */

const REACHABILITY_PROBE_TIMEOUT_MS = 5000;

class ApiError extends Error {
  constructor(message, { status = 0, body = null, cause = null } = {}) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.body = body;
    if (cause) {
      this.cause = cause;
    }
  }
}

function encodeUtf8Base64(value, btoaImplementation = globalThis.btoa) {
  if (typeof btoaImplementation !== "function") {
    throw new Error("This browser does not support Basic authentication encoding.");
  }
  const bytes = encodeURIComponent(value).replace(/%([0-9A-F]{2})/g, (_match, hex) =>
    String.fromCharCode(Number.parseInt(hex, 16))
  );
  return btoaImplementation(bytes);
}

function createAuthorization(auth) {
  const mode = auth?.mode;
  if (mode === "bearer") {
    const token = auth.token?.trim();
    if (!token) {
      throw new Error("Enter an access token.");
    }
    return `Bearer ${token}`;
  }
  if (mode === "basic") {
    if (!auth.username || !auth.password) {
      throw new Error("Enter a username and password.");
    }
    return `Basic ${encodeUtf8Base64(`${auth.username}:${auth.password}`)}`;
  }
  throw new Error(`Unsupported authentication mode: ${mode || "empty"}.`);
}

function extractErrorMessage(body, fallback) {
  if (!body) {
    return fallback;
  }
  if (typeof body === "string") {
    return body.trim() || fallback;
  }
  const details = body["o:errorDetails"];
  if (Array.isArray(details) && details.length > 0) {
    const detail = details.find((item) => item && (item.detail || item.title || item.message));
    if (detail) {
      return detail.detail || detail.title || detail.message;
    }
  }
  return body.detail || body.title || body.message || fallback;
}

async function parseResponse(response) {
  if (response.status === 204) {
    return null;
  }
  const text = await response.text();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function extractResourceKey(item, environmentUrl) {
  const links = Array.isArray(item?.links) ? item.links : [];
  const link =
    links.find((candidate) => candidate?.rel === "self") ||
    links.find((candidate) => candidate?.rel === "canonical");
  if (!link?.href) {
    throw new ApiError(
      "A revenue item did not include its Oracle resource link. Ensure the configured API or gateway preserves item links."
    );
  }

  let pathname;
  try {
    pathname = new URL(link.href, environmentUrl).pathname;
  } catch {
    throw new ApiError("A revenue item returned an invalid Oracle resource link.", { body: link });
  }
  const encodedKey = pathname.split("/").filter(Boolean).at(-1);
  let key;
  try {
    key = decodeURIComponent(encodedKey || "");
  } catch {
    throw new ApiError("A revenue item returned an invalid encoded resource key.", { body: link });
  }
  if (!key || key === "." || key === "..") {
    throw new ApiError("A revenue item returned an empty or unsafe resource key.", { body: link });
  }
  return key;
}

class ApiClient {
  constructor(config, auth, fetchImplementation = fetch) {
    this.config = config;
    this.fetch = fetchImplementation;
    this.authorization = createAuthorization(auth);
  }

  createUrl(resourcePath = "", query = {}) {
    const url = new URL(
      `${this.config.api.environmentUrl}${this.config.api.restPath}${resourcePath}`
    );
    for (const [name, value] of Object.entries(query)) {
      if (value !== undefined && value !== null && value !== "") {
        url.searchParams.set(name, String(value));
      }
    }
    return url.toString();
  }

  async probeReachability(url) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), REACHABILITY_PROBE_TIMEOUT_MS);
    try {
      await this.fetch(url, {
        credentials: "omit",
        method: "GET",
        mode: "no-cors",
        signal: controller.signal,
      });
      return true;
    } catch {
      return false;
    } finally {
      clearTimeout(timeout);
    }
  }

  async request(resourcePath, { method = "GET", query, body, contentType } = {}) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.config.api.requestTimeoutMs);
    const headers = {
      Accept: "application/json",
      "REST-Framework-Version": this.config.api.frameworkVersion,
    };
    headers.Authorization = this.authorization;
    if (body !== undefined) {
      headers["Content-Type"] = contentType || "application/json";
    }

    const requestUrl = this.createUrl(resourcePath, query);
    let response;
    try {
      response = await this.fetch(requestUrl, {
        body: body === undefined ? undefined : JSON.stringify(body),
        credentials: "omit",
        headers,
        method,
        signal: controller.signal,
      });
    } catch (error) {
      if (error && error.name === "AbortError") {
        throw new ApiError(`Request timed out after ${this.config.api.requestTimeoutMs} ms.`, {
          cause: error,
        });
      }
      if (method === "GET" && body === undefined && (await this.probeReachability(requestUrl))) {
        throw new ApiError(
          "The service appears reachable, but the browser blocked the authenticated REST request. Configure CORS for the add-in origin and allow the Authorization and REST-Framework-Version headers.",
          { cause: error }
        );
      }
      throw new ApiError(
        "Could not verify network access to the environment. Check the environment URL, DNS or VPN access, and its TLS certificate. CORS may also be blocking the REST request.",
        { cause: error }
      );
    } finally {
      clearTimeout(timeout);
    }

    const responseBody = await parseResponse(response);
    if (!response.ok) {
      const fallback =
        response.status === 401
          ? "Authentication failed or the access token is invalid or expired."
          : response.status === 403
            ? "The signed-in user is not authorized for this operation."
            : `Service request failed (${response.status}).`;
      throw new ApiError(extractErrorMessage(responseBody, fallback), {
        body: responseBody,
        status: response.status,
      });
    }

    return responseBody;
  }

  async getCollection(resourcePath, query, maximumItems, { onlyData = true } = {}) {
    const items = [];
    let offset = 0;

    while (items.length < maximumItems) {
      const limit = Math.min(this.config.api.pageSize, maximumItems - items.length);
      const page = await this.request(resourcePath, {
        query: { ...query, limit, offset, onlyData },
      });
      if (!page || !Array.isArray(page.items)) {
        throw new ApiError("The service returned an invalid collection response.", { body: page });
      }
      items.push(...page.items);
      if (page.hasMore !== true || page.items.length === 0) {
        break;
      }
      offset += page.items.length;
    }

    return items;
  }

  async testConnection() {
    return this.getCollection(this.config.resources.opportunities, { fields: "OptyNumber" }, 1);
  }

  async getOpportunities() {
    return this.getCollection(
      this.config.resources.opportunities,
      {
        fields:
          "OptyId,OptyNumber,Name,TargetPartyName,Revenue,CurrencyCode,SalesStage,WinProb,EffectiveDate",
        orderBy: "EffectiveDate:desc",
      },
      this.config.api.opportunityLimit
    );
  }

  async getProductGroups() {
    const idField = this.config.resources.productGroupIdField;
    const nameField = this.config.resources.productGroupNameField;
    const items = await this.getCollection(
      this.config.resources.productGroups,
      { fields: `${idField},${nameField}`, orderBy: nameField },
      this.config.api.productGroupLimit
    );
    return items.map((item) => ({
      ...item,
      ProdGroupId: item[idField],
      ProdGroupName: item[nameField],
    }));
  }

  revenueItemsPath(opportunityNumber) {
    return this.config.resources.revenueItems.replace(
      "{opportunityNumber}",
      encodeURIComponent(String(opportunityNumber))
    );
  }

  async getRevenueItems(opportunityNumber) {
    const items = await this.getCollection(
      this.revenueItemsPath(opportunityNumber),
      { fields: "RevnId,ProductType,ProdGroupId,ProdGroupName,Quantity,UnitPrice,RevnAmount" },
      this.config.api.revenueItemLimit,
      { onlyData: false }
    );
    return items.map((item) => ({
      ...item,
      _resourceKey: extractResourceKey(item, this.config.api.environmentUrl),
    }));
  }

  async syncRevenueItems(parts) {
    const response = await this.request("", {
      body: { parts },
      contentType: "application/vnd.oracle.adf.batch+json",
      method: "POST",
    });
    if (
      response !== null &&
      (!Array.isArray(response.parts) || response.parts.length !== parts.length)
    ) {
      throw new ApiError("The service returned an incomplete or invalid batch response.", {
        body: response,
      });
    }
    return response;
  }
}

module.exports = {
  ApiClient,
  ApiError,
  createAuthorization,
  encodeUtf8Base64,
  extractResourceKey,
  extractErrorMessage,
};
