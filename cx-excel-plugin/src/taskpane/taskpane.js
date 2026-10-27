/*
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
*/

/* global console, document, module, require, window */

const { ApiClient, ApiError } = require("./api-client");
const { loadRuntimeConfig, normalizeConfig } = require("./config");

const HEADERS = [
  "Resource key",
  "Type",
  "Product group",
  "Quantity",
  "Unit price",
  "Amount",
  "Message",
];

const COLUMN = {
  id: 0,
  type: 1,
  name: 2,
  quantity: 3,
  unitPrice: 4,
  amount: 5,
  message: 6,
};

const state = {
  busy: false,
  client: null,
  config: null,
  loadedOpportunityNumber: "",
  opportunities: new Map(),
  originals: new Map(),
  productGroupIdsByName: new Map(),
  productGroupLabelsById: new Map(),
  productGroupNames: [],
};

const element = (id) => document.getElementById(id);

Office.onReady(async (info) => {
  element("app-body").hidden = false;
  bindEvents();

  if (info.host !== Office.HostType.Excel) {
    showConnectionError("This add-in must be opened in Excel.");
    return;
  }
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.8")) {
    showConnectionError("This version of Excel does not support the required ExcelApi 1.8 APIs.");
    return;
  }

  try {
    state.config = await loadRuntimeConfig();
    applyRuntimeConfig();
    showConnection();
  } catch (error) {
    showConnection();
    showConnectionError("Configuration error: " + error.message);
  }
});

function bindEvents() {
  element("connection-form").addEventListener("submit", connect);
  element("auth-mode").addEventListener("change", updateAuthenticationFields);
  element("reset-button").addEventListener("click", resetConnectionSettings);
  element("disconnect-button").addEventListener("click", disconnect);
  element("opportunity-select").addEventListener("change", selectOpportunity);
  element("load-button").addEventListener("click", loadRevenueItems);
  element("sync-button").addEventListener("click", syncRevenueItems);
  element("view-button").addEventListener("click", viewOpportunity);
}

function applyRuntimeConfig() {
  document.title = state.config.ui.title;
  element("app-title").textContent = state.config.ui.title;
  element("brand-logo").src = state.config.ui.logoPath;
  element("environment-url").value = state.config.api.environmentUrl;
  element("auth-mode").value = state.config.auth.mode;
  element("auth-mode").disabled = !state.config.auth.allowModeSelection;
  element("debug-section").hidden = !state.config.ui.showDebug;
  element("view-button").hidden = !state.config.ui.opportunityUrlTemplate;
  updateAuthenticationFields();
}

function resetConnectionSettings() {
  clearSecrets();
  applyRuntimeConfig();
  showConnectionError("");
}

function clearSecrets() {
  element("username").value = "";
  element("password").value = "";
  element("access-token").value = "";
}

function updateAuthenticationFields() {
  const mode = element("auth-mode").value;
  element("basic-credentials").hidden = mode !== "basic";
  element("bearer-credentials").hidden = mode !== "bearer";
}

function readAuthentication() {
  const mode = element("auth-mode").value;
  return {
    mode,
    password: mode === "basic" ? element("password").value : "",
    token: mode === "bearer" ? element("access-token").value : "",
    username: mode === "basic" ? element("username").value.trim() : "",
  };
}

function connectionConfig() {
  return normalizeConfig({
    ...state.config,
    api: {
      ...state.config.api,
      environmentUrl: element("environment-url").value,
    },
    auth: {
      ...state.config.auth,
      mode: element("auth-mode").value,
    },
  });
}

async function connect(event) {
  event.preventDefault();
  showConnectionError("");
  showLoader("Checking the connection…");

  try {
    const config = connectionConfig();
    if (!config.api.environmentUrl) {
      throw new Error("Enter an environment URL.");
    }
    const client = new ApiClient(config, readAuthentication());
    await client.testConnection();
    element("loader-message").textContent = "Loading opportunities and product groups…";
    const [opportunities, productGroups] = await Promise.all([
      client.getOpportunities(),
      client.getProductGroups(),
    ]);

    state.client = client;
    state.config = config;
    state.loadedOpportunityNumber = "";
    state.originals.clear();
    setOpportunities(opportunities);
    setProductGroups(productGroups);
    clearSecrets();
    showRevenue();
    setStatus(
      opportunities.length
        ? "Connected. Select an opportunity."
        : "Connected, but no accessible opportunities were returned."
    );
  } catch (error) {
    state.client = null;
    showConnection();
    showConnectionError(error.message || "Unable to connect.");
    showDebug(error);
  }
}

function disconnect() {
  state.client = null;
  state.loadedOpportunityNumber = "";
  state.originals.clear();
  clearSecrets();
  showConnection();
  showConnectionError("");
}

function setOpportunities(opportunities) {
  state.opportunities.clear();
  const select = element("opportunity-select");
  select.replaceChildren();
  addSelectOption(select, "Select an opportunity", "");

  for (const opportunity of opportunities) {
    const number = String(opportunity.OptyNumber || "");
    if (!number) {
      continue;
    }
    state.opportunities.set(number, opportunity);
    addSelectOption(select, (opportunity.Name || "Unnamed") + " — " + number, number);
  }

  select.disabled = false;
  select.value = "";
  refreshControls();
}

function addSelectOption(select, text, value) {
  const option = document.createElement("option");
  option.textContent = text;
  option.value = value;
  select.append(option);
}

function setProductGroups(productGroups) {
  state.productGroupIdsByName.clear();
  state.productGroupLabelsById.clear();
  state.productGroupNames = [];

  for (const productGroup of productGroups) {
    const name = String(productGroup.ProdGroupName || "").trim();
    if (!name || productGroup.ProdGroupId === undefined || productGroup.ProdGroupId === null) {
      continue;
    }
    const id = String(productGroup.ProdGroupId);
    const label = `${name} [${id}]`;
    state.productGroupNames.push(label);
    state.productGroupIdsByName.set(label, productGroup.ProdGroupId);
    state.productGroupLabelsById.set(id, label);
  }
}

function prepareRevenueItems(items) {
  return items.map((item) => ({
    ...item,
    _productGroupLabel:
      state.productGroupLabelsById.get(String(item.ProdGroupId)) || item.ProdGroupName || "",
  }));
}

function selectedOpportunity() {
  return state.opportunities.get(element("opportunity-select").value) || null;
}

function selectOpportunity() {
  state.loadedOpportunityNumber = "";
  state.originals.clear();
  const opportunity = selectedOpportunity();
  renderOpportunity(opportunity);
  refreshControls();
  setStatus(opportunity ? "Load the revenue items into Excel to begin editing." : "");
}

function renderOpportunity(opportunity) {
  element("opportunity-account").textContent = opportunity?.TargetPartyName || "—";
  element("opportunity-amount").textContent = opportunity
    ? formatAmount(opportunity.Revenue, opportunity.CurrencyCode)
    : "—";
  element("opportunity-sales-stage").textContent = opportunity?.SalesStage || "—";
  element("opportunity-win-probability").textContent =
    opportunity?.WinProb === undefined || opportunity?.WinProb === null
      ? "—"
      : String(opportunity.WinProb) + "%";
  element("opportunity-close-date").textContent = opportunity?.EffectiveDate || "—";
}

function formatAmount(amount, currencyCode) {
  const value = Number(amount);
  if (!Number.isFinite(value)) {
    return amount || "—";
  }
  try {
    return new Intl.NumberFormat(undefined, {
      currency: currencyCode || undefined,
      maximumFractionDigits: 2,
      minimumFractionDigits: 2,
      style: currencyCode ? "currency" : "decimal",
    }).format(value);
  } catch {
    return value.toFixed(2) + (currencyCode ? " " + currencyCode : "");
  }
}

function refreshControls() {
  const opportunity = selectedOpportunity();
  const selectedNumber = opportunity ? String(opportunity.OptyNumber) : "";
  element("opportunity-select").disabled = !state.client || state.busy;
  element("load-button").disabled = !state.client || !opportunity || state.busy;
  element("sync-button").disabled =
    !state.client || !opportunity || state.busy || state.loadedOpportunityNumber !== selectedNumber;
  element("view-button").disabled = !opportunity || state.busy;
}

function setBusy(busy) {
  state.busy = busy;
  refreshControls();
}

async function loadRevenueItems() {
  const opportunity = selectedOpportunity();
  if (!opportunity) {
    setStatus("Select an opportunity first.", true);
    return;
  }

  setBusy(true);
  setStatus("Loading revenue items…");
  try {
    const items = prepareRevenueItems(await state.client.getRevenueItems(opportunity.OptyNumber));
    showDebug(items);
    await writeRevenueWorksheet(items);
    state.originals = new Map(items.map((item) => [item._resourceKey, item]));
    state.loadedOpportunityNumber = String(opportunity.OptyNumber);
    setStatus("Loaded " + items.length + " revenue item" + (items.length === 1 ? "." : "s."));
  } catch (error) {
    handleOperationError(error);
  } finally {
    setBusy(false);
  }
}

async function writeRevenueWorksheet(items) {
  if (items.length > state.config.workbook.maxRows) {
    throw new Error(
      "The service returned " +
        items.length +
        " items, but workbook.maxRows is " +
        state.config.workbook.maxRows +
        "."
    );
  }

  await Excel.run(async (context) => {
    const worksheet = await getOrCreateWorksheet(context);
    const maximumRow = state.config.workbook.maxRows + 1;
    const managedRange = worksheet.getRange("A1:G" + maximumRow);
    managedRange.clear(Excel.ClearApplyTo.all);

    const headerRange = worksheet.getRange("A1:G1");
    headerRange.values = [HEADERS];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#E8E2DE";
    headerRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

    if (items.length) {
      const values = items.map((item) => [
        item._resourceKey,
        item.ProductType || "Group",
        item._productGroupLabel,
        item.Quantity ?? "",
        item.UnitPrice ?? "",
        item.RevnAmount ?? "",
        "",
      ]);
      worksheet.getRange("A2:G" + (items.length + 1)).values = values;
    }

    const editableRows = state.config.workbook.maxRows;
    worksheet.getRange("A2:A" + maximumRow).numberFormat = Array.from(
      { length: editableRows },
      () => ["@"]
    );
    const numericFormat = Array.from({ length: editableRows }, () => [
      state.config.workbook.numberFormat,
    ]);
    worksheet.getRange("E2:E" + maximumRow).numberFormat = numericFormat;
    worksheet.getRange("F2:F" + maximumRow).numberFormat = numericFormat;

    worksheet.getRange("B2:B" + maximumRow).dataValidation.rule = {
      list: { inCellDropDown: true, source: "Group" },
    };
    const inlineProductGroups = state.productGroupNames.join(",");
    if (
      state.productGroupNames.length &&
      inlineProductGroups.length <= 255 &&
      !state.productGroupNames.some((name) => name.includes(","))
    ) {
      worksheet.getRange("C2:C" + maximumRow).dataValidation.rule = {
        list: { inCellDropDown: true, source: inlineProductGroups },
      };
    }

    worksheet.getRange("A:A").format.columnWidth = 180;
    worksheet.getRange("B:B").format.columnWidth = 70;
    worksheet.getRange("C:C").format.columnWidth = 180;
    worksheet.getRange("D:F").format.columnWidth = 90;
    worksheet.getRange("G:G").format.columnWidth = 220;
    worksheet.getRange("A1:G" + maximumRow).format.verticalAlignment =
      Excel.VerticalAlignment.center;
    worksheet.getRange("A1:G" + maximumRow).format.autofitRows();
    worksheet.activate();
    await context.sync();
  });
}

async function getOrCreateWorksheet(context) {
  const worksheets = context.workbook.worksheets;
  let worksheet = worksheets.getItemOrNullObject(state.config.workbook.worksheetName);
  // Loading this scalar is required before checking an OrNullObject result.
  // eslint-disable-next-line office-addins/no-navigational-load
  worksheet.load("isNullObject");
  await context.sync();
  if (worksheet.isNullObject) {
    worksheet = worksheets.add(state.config.workbook.worksheetName);
  }
  return worksheet;
}

async function readRevenueRows() {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(state.config.workbook.worksheetName);
    const range = worksheet.getRange("A2:G" + (state.config.workbook.maxRows + 1));
    range.load("values");
    await context.sync();
    return range.values;
  });
}

function textValue(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function numberValue(value, label, rowNumber) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new Error(`Row ${rowNumber}: ${label} must be a number.`);
  }
  return number;
}

function sameValue(left, right) {
  if (
    (left === "" || left === null || left === undefined) &&
    (right === "" || right === null || right === undefined)
  ) {
    return true;
  }
  const leftNumber = Number(left);
  const rightNumber = Number(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return leftNumber === rightNumber;
  }
  return textValue(left) === textValue(right);
}

function rowChanged(record, original) {
  return (
    !sameValue(record.productType, original.ProductType) ||
    !sameValue(record.productGroupName, original._productGroupLabel || original.ProdGroupName) ||
    !sameValue(record.quantity, original.Quantity) ||
    !sameValue(record.unitPrice, original.UnitPrice) ||
    !sameValue(record.amount, original.RevnAmount)
  );
}

function readRow(row, rowNumber) {
  return {
    amount: numberValue(row[COLUMN.amount], "Amount", rowNumber),
    id: textValue(row[COLUMN.id]),
    productGroupName: textValue(row[COLUMN.name]),
    productType: textValue(row[COLUMN.type]) || "Group",
    quantity: numberValue(row[COLUMN.quantity], "Quantity", rowNumber),
    unitPrice: numberValue(row[COLUMN.unitPrice], "Unit price", rowNumber),
  };
}

function createPayload(record, productGroupId, extraFields, includeNulls) {
  const payload = {
    ...extraFields,
    ProductType: "Group",
    ProdGroupId: productGroupId,
  };
  if (includeNulls || record.quantity !== null) {
    payload.Quantity = record.quantity;
  }
  if (includeNulls || record.unitPrice !== null) {
    payload.UnitPrice = record.unitPrice;
  }
  if (includeNulls || record.amount !== null) {
    payload.RevnAmount = record.amount;
  }
  return payload;
}

function buildBatchParts({
  rows,
  originals,
  productGroupIdsByName,
  revenueItemsPath,
  extraFields = {},
}) {
  const parts = [];
  const rowByPartId = new Map();

  rows.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + 2;
    const record = readRow(row, rowNumber);
    const original = record.id ? originals.get(record.id) : null;
    const completelyEmpty =
      !record.id &&
      !record.productGroupName &&
      record.quantity === null &&
      record.unitPrice === null &&
      record.amount === null;

    if (completelyEmpty) {
      return;
    }
    if (record.productType !== "Group") {
      if (original && !rowChanged(record, original)) {
        return;
      }
      throw new Error(`Row ${rowNumber}: only product-group revenue items are supported.`);
    }
    if (!record.productGroupName) {
      throw new Error(`Row ${rowNumber}: select a product group.`);
    }

    let productGroupId = productGroupIdsByName.get(record.productGroupName);
    const originalProductGroupName =
      original && (original._productGroupLabel || original.ProdGroupName);
    if (
      (productGroupId === undefined || productGroupId === null) &&
      original &&
      sameValue(record.productGroupName, originalProductGroupName)
    ) {
      productGroupId = original.ProdGroupId;
    }
    if (productGroupId === undefined || productGroupId === null) {
      throw new Error(
        `Row ${rowNumber}: product group "${record.productGroupName}" was not found in the configured catalog.`
      );
    }

    const isNew = !record.id || record.id === "0";
    if (!isNew && !original) {
      throw new Error(
        `Row ${rowNumber}: load revenue items before updating resource key ${record.id}.`
      );
    }
    if (!isNew && !rowChanged(record, original)) {
      return;
    }

    const partId = `part${parts.length + 1}`;
    parts.push({
      id: partId,
      operation: isNew ? "create" : "update",
      path: isNew ? revenueItemsPath : `${revenueItemsPath}/${encodeURIComponent(record.id)}`,
      payload: createPayload(record, productGroupId, extraFields, !isNew),
    });
    rowByPartId.set(partId, rowIndex);
  });

  return { parts, rowByPartId };
}

async function writeRowErrors(rowByPartId, message) {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(state.config.workbook.worksheetName);
    const messageRange = worksheet.getRange("G2:G" + (state.config.workbook.maxRows + 1));
    messageRange.clear(Excel.ClearApplyTo.contents);
    for (const rowIndex of rowByPartId.values()) {
      const cell = messageRange.getCell(rowIndex, 0);
      cell.values = [[message]];
      cell.format.font.color = "#B42318";
    }
    await context.sync();
  });
}

async function syncRevenueItems() {
  const opportunity = selectedOpportunity();
  if (!opportunity || state.loadedOpportunityNumber !== String(opportunity.OptyNumber)) {
    setStatus("Load this opportunity before syncing.", true);
    return;
  }

  let batch;
  setBusy(true);
  setStatus("Reading workbook changes…");
  try {
    const rows = await readRevenueRows();
    batch = buildBatchParts({
      extraFields: state.config.sync.extraRevenueFields,
      originals: state.originals,
      productGroupIdsByName: state.productGroupIdsByName,
      revenueItemsPath: state.client.revenueItemsPath(opportunity.OptyNumber),
      rows,
    });
    if (!batch.parts.length) {
      setStatus("No changes to sync.");
      return;
    }

    setStatus(
      "Syncing " + batch.parts.length + " change" + (batch.parts.length === 1 ? "…" : "s…")
    );
    const response = await state.client.syncRevenueItems(batch.parts);
    showDebug(response);
    const refreshedItems = prepareRevenueItems(
      await state.client.getRevenueItems(opportunity.OptyNumber)
    );
    await writeRevenueWorksheet(refreshedItems);
    state.originals = new Map(refreshedItems.map((item) => [item._resourceKey, item]));
    setStatus("Synced " + batch.parts.length + " change" + (batch.parts.length === 1 ? "." : "s."));
  } catch (error) {
    if (batch?.rowByPartId?.size) {
      try {
        await writeRowErrors(batch.rowByPartId, error.message || "Not synced");
      } catch (worksheetError) {
        console.error("Could not write row errors.", worksheetError);
      }
    }
    handleOperationError(error);
  } finally {
    setBusy(false);
  }
}

function viewOpportunity() {
  const opportunity = selectedOpportunity();
  const template = state.config.ui.opportunityUrlTemplate;
  if (!opportunity || !template) {
    return;
  }
  const url = template
    .replaceAll("{opportunityId}", encodeURIComponent(String(opportunity.OptyId || "")))
    .replaceAll("{opportunityNumber}", encodeURIComponent(String(opportunity.OptyNumber || "")));
  window.open(url, "_blank", "noopener");
}

function handleOperationError(error) {
  showDebug(error);
  if (error instanceof ApiError && error.status === 401) {
    disconnect();
    showConnectionError(error.message);
    return;
  }
  setStatus(error.message || "The operation failed.", true);
}

function showConnection() {
  element("connection-section").hidden = false;
  element("loader-section").hidden = true;
  element("revenue-section").hidden = true;
}

function showLoader(message) {
  element("connection-section").hidden = true;
  element("loader-section").hidden = false;
  element("revenue-section").hidden = true;
  element("loader-message").textContent = message;
}

function showRevenue() {
  element("connection-section").hidden = true;
  element("loader-section").hidden = true;
  element("revenue-section").hidden = false;
  renderOpportunity(null);
  refreshControls();
}

function showConnectionError(message) {
  element("connection-error").textContent = message;
}

function setStatus(message, error = false) {
  const status = element("status-message");
  status.textContent = message;
  status.classList.toggle("error", error);
}

function showDebug(value) {
  if (!state.config?.ui.showDebug) {
    return;
  }
  const serializable =
    value instanceof Error
      ? { body: value.body, message: value.message, name: value.name, status: value.status }
      : value;
  element("debug-response").value = JSON.stringify(serializable, null, 2);
}

module.exports = { buildBatchParts };
