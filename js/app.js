const appState = {
  catalogByCode: new Map(),
  puntadasByTipoMaq: new Map(),
  costuraProtectionValue: APP_CONFIG.defaultProtection,
  corteProtectionValue: APP_CONFIG.defaultProtection,
  acabadoProtectionValue: APP_CONFIG.defaultProtection,
  searchResults: [],
  activeRecordKey: null,
  masterHeaders: [],
  masterSourceRows: [],
  masterRows: [],
  masterIsLoading: false,
  masterFilters: {},
  masterEditingRowKey: null,
  currentView: "editor",
  interactionMode: "create",
  lastSearchProto: "",
};

const DEFAULT_CLIENT_OPTIONS = Object.freeze([
  "LULULEMON",
  "BANANA",
  "SKECHERS",
  "ATHLETA",
  "THEORY",
  "DUER",
  "ALLBIRDS",
  "LACOSTE",
  "AM RETAIL",
]);

const NEW_CLIENT_OPTION_VALUE = "__NEW_CLIENT__";
const CUSTOM_CLIENTS_STORAGE_KEY = "costura_custom_clients";
const EDITABLE_MODES = new Set(["create", "search_edit", "search_new_version"]);
const VERSION_COLLATOR = new Intl.Collator("es-PE", { numeric: true, sensitivity: "base" });
const MASTER_EDIT_TRIGGER_HEADER = "DESCRIPCION_ABREVIADA";
const MASTER_DATE_HEADER = "FECHA_REGISTRO";
const MASTER_TEXT_FILTER_HEADERS = new Set(["CODIGO", "DESCRIPCION_ABREVIADA", "DESCRIPCION_COMPLETA"]);
const MASTER_SELECT_FILTER_HEADERS = new Set(["BLOQUE", "TIPO_OPERACION", "ACCION", "TIPO_MAQ", "MAQUINA"]);
const MASTER_LOADING_MESSAGE = "Cargando operaciones...";

const dom = {};
let toastTimerId = null;
let confirmAction = null;
let codeContextMenuRow = null;
let codeContextMenuTable = "costura"; // tracks origin table: 'costura', 'corte', or 'acabado'
let masterDraftSequence = 0;

document.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  cacheDom();
  showLoadingModal("Cargando datos");
  initializeClientField();
  bindEvents();
  buildInitialRows();
  refreshSummary();
  setInteractionMode("create");
  switchView("editor");
  updateMasterActionState();

  if (!AppsScriptAPI.isConfigured()) {
    setConnectionState("pending", "Configura WEB_APP_URL");
    showToast("Actualiza js/config.js con la URL del Web App de Apps Script para activar la conexiÃ³n.", "info");
    hideLoadingModal();
    return;
  }

  setConnectionState("pending", "Conectando a Sheets...");

  try {
    await refreshCatalogData({ notifySuccess: false });
  } catch (error) {
    setConnectionState("offline", "Sin conexiÃ³n con Sheets");
    showToast(error.message || "No se pudo cargar el catÃ¡logo desde Google Sheets.", "error");
  } finally {
    hideLoadingModal();
  }
}

async function refreshCatalogData(options = {}) {
  const { notifySuccess = true } = options;
  const response = await AppsScriptAPI.fetchCatalog();
  const payload = response.data || response;

  if (!response.success && !payload.basedatos) {
    throw new Error(response.message || "No se pudo cargar el catÃ¡logo.");
  }

  hydrateCatalog(payload);
  setConnectionState("online", `CatÃ¡logo cargado: ${appState.catalogByCode.size} cÃ³digos`);

  if (notifySuccess) {
    showToast("CatÃ¡logo de costura listo para usar.", "success");
  }

  return payload;
}

function cacheDom() {
  dom.appShell = document.querySelector(".app-shell");
  dom.newPrototypeBtn = document.getElementById("newPrototypeBtn");
  dom.openSearchBtn = document.getElementById("openSearchBtn");
  dom.openMasterBtn = document.getElementById("openMasterBtn");
  dom.searchVersionGrid = document.querySelector(".search-version-grid");
  dom.searchPanel = document.getElementById("searchPanel");
  dom.searchProtoInput = document.getElementById("searchProtoInput");
  dom.searchBtn = document.getElementById("searchBtn");
  dom.searchSummary = document.getElementById("searchSummary");
  dom.versionPanel = document.getElementById("versionPanel");
  dom.versionTabs = document.getElementById("versionTabs");
  dom.versionInfo = document.getElementById("versionInfo");
  dom.versionSubinfo = document.getElementById("versionSubinfo");
  dom.editVersionBtn = document.getElementById("editVersionBtn");
  dom.newVersionBtn = document.getElementById("newVersionBtn");
  dom.printRecordBtn = document.getElementById("printRecordBtn");
  dom.connectionBadge = document.getElementById("connectionBadge");
  dom.formPanel = document.getElementById("formPanel");
  dom.formRecordInfo = document.getElementById("formRecordInfo");
  dom.costuraPanel = document.getElementById("costuraPanel");
  dom.cortePanel = document.getElementById("cortePanel");
  dom.acabadoPanel = document.getElementById("acabadoPanel");
  dom.resumenPanel = document.getElementById("resumenPanel");
  dom.tabCostura = document.getElementById("tabCostura");
  dom.tabCorte = document.getElementById("tabCorte");
  dom.tabAcabado = document.getElementById("tabAcabado");
  dom.tabResumen = document.getElementById("tabResumen");
  dom.panelTabs = document.querySelector(".panel-tabs");
  dom.masterPanel = document.getElementById("masterPanel");
  dom.masterInfo = document.getElementById("masterInfo");
  dom.masterSaveBtn = document.getElementById("masterSaveBtn");
  dom.masterExcelBtn = document.getElementById("masterExcelBtn");
  dom.masterAddBtn = document.getElementById("masterAddBtn");
  dom.masterTableColgroup = document.getElementById("masterTableColgroup");
  dom.masterTableHead = document.getElementById("masterTableHead");
  dom.masterTableBody = document.getElementById("masterTableBody");
  dom.masterFilterMenu = document.getElementById("masterFilterMenu");
  dom.masterFilterTitle = document.getElementById("masterFilterTitle");
  dom.masterFilterLabel = document.getElementById("masterFilterLabel");
  dom.masterFilterTextInput = document.getElementById("masterFilterTextInput");
  dom.masterFilterSelect = document.getElementById("masterFilterSelect");
  dom.masterFilterClearBtn = document.getElementById("masterFilterClearBtn");
  dom.saveBtn = document.getElementById("saveBtn");
  dom.copyPrototypeBtn = document.getElementById("copyPrototypeBtn");
  dom.clearPrototypeBtn = document.getElementById("clearPrototypeBtn");
  dom.addRowBtn = document.getElementById("addRowBtn");
  dom.removeRowBtn = document.getElementById("removeRowBtn");
  dom.clearTableBtn = document.getElementById("clearTableBtn");
  dom.addCorteRowBtn = document.getElementById("addCorteRowBtn");
  dom.removeCorteRowBtn = document.getElementById("removeCorteRowBtn");
  dom.clearCorteTableBtn = document.getElementById("clearCorteTableBtn");
  dom.addAcabadoRowBtn = document.getElementById("addAcabadoRowBtn");
  dom.removeAcabadoRowBtn = document.getElementById("removeAcabadoRowBtn");
  dom.clearAcabadoTableBtn = document.getElementById("clearAcabadoTableBtn");
  
  dom.tableBody = document.getElementById("costuraTableBody");
  dom.rowTemplate = document.getElementById("costuraRowTemplate");
  dom.corteTableBody = document.getElementById("corteTableBody");
  dom.corteRowTemplate = document.getElementById("corteRowTemplate");
  dom.acabadoTableBody = document.getElementById("acabadoTableBody");
  dom.acabadoRowTemplate = document.getElementById("acabadoRowTemplate");
  
  dom.codigoSuggestions = document.getElementById("codigoSuggestions");
  dom.toast = document.getElementById("toast");
  dom.loadingModal = document.getElementById("loadingModal");
  dom.loadingModalCard = document.getElementById("loadingModalCard");
  dom.loadingModalTitle = document.getElementById("loadingModalTitle");
  dom.confirmModal = document.getElementById("confirmModal");
  dom.confirmModalTitle = document.getElementById("confirmModalTitle");
  dom.confirmModalMessage = document.getElementById("confirmModalMessage");
  dom.confirmModalCancelBtn = document.getElementById("confirmModalCancelBtn");
  dom.confirmModalAcceptBtn = document.getElementById("confirmModalAcceptBtn");
  dom.copyPrototypeModal = document.getElementById("copyPrototypeModal");
  dom.copyPrototypeInput = document.getElementById("copyPrototypeInput");
  dom.copyPrototypeCancelBtn = document.getElementById("copyPrototypeCancelBtn");
  dom.copyPrototypeAcceptBtn = document.getElementById("copyPrototypeAcceptBtn");
  dom.codeContextMenu = document.getElementById("codeContextMenu");

  dom.form = {
    cliente: document.getElementById("clienteSelect"),
    clienteDelete: document.getElementById("deleteClientBtn"),
    clienteCustom: document.getElementById("clienteCustomInput"),
    proto: document.getElementById("protoInput"),
    version: document.getElementById("versionInput"),
    idem: document.getElementById("idemInput"),
    descripcion: document.getElementById("descripcionInput"),
    estilo: document.getElementById("estiloInput"),
    tela: document.getElementById("telaInput"),
    realizadoPor: document.getElementById("realizadoPorInput"),
    produccionEstimada: document.getElementById("produccionEstimadaInput"),
    rutasProcesos: document.getElementById("rutasProcesosInput"),
  };

  dom.summary = {
    estimatedFooter: document.getElementById("estimatedFooterValue"),
    maqFooter: document.getElementById("maqFooterValue"),
    manualFooter: document.getElementById("manualFooterValue"),
    cotizacionFooter: document.getElementById("cotizacionFooterValue"),
  };
  dom.summaryCorte = {
    estimatedCorte: document.getElementById("estimatedCorteFooterValue"),
    estimatedHab: document.getElementById("estimatedHabFooterValue"),
    corte: document.getElementById("corteFooterValue"),
    hab: document.getElementById("habFooterValue"),
    cotizacion: document.getElementById("cotizacionCorteFooterValue"),
  };
  dom.summaryAcabado = {
    estimated: document.getElementById("estimatedAcabadoFooterValue"),
    cotizacion: document.getElementById("cotizacionAcabadoFooterValue"),
  };
  dom.summaryResumen = {
    corte: document.getElementById("resumenCorteTotalValue"),
    costuraMaq: document.getElementById("resumenCosturaMaqValue"),
    costuraManual: document.getElementById("resumenCosturaManualValue"),
    costuraTotal: document.getElementById("resumenCosturaTotalValue"),
    acabados: document.getElementById("resumenAcabadosValue"),
    total: document.getElementById("resumenGranTotalValue"),
  };
}

function bindEvents() {
  dom.newPrototypeBtn.addEventListener("click", handleNewPrototype);
  dom.openSearchBtn.addEventListener("click", () => switchView("search"));
  dom.openMasterBtn.addEventListener("click", handleOpenMasterView);
  dom.searchBtn.addEventListener("click", handleSearch);
  dom.editVersionBtn.addEventListener("click", handleEditVersion);
  dom.newVersionBtn.addEventListener("click", handleNewVersionFromRecord);
  dom.printRecordBtn.addEventListener("click", handlePrintRecord);
  dom.form.cliente.addEventListener("change", handleClientSelectChange);
  dom.form.clienteDelete.addEventListener("click", deleteSelectedCustomClient);
  dom.form.clienteCustom.addEventListener("change", commitCustomClient);
  dom.form.clienteCustom.addEventListener("blur", commitCustomClient);
  dom.form.clienteCustom.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      commitCustomClient();
    }
  });
  dom.form.proto.addEventListener("input", handleProtoInput);
  dom.form.version.addEventListener("input", handleVersionInput);
  dom.form.idem.addEventListener("input", handleIdemInput);
  [
    dom.form.clienteCustom,
    dom.form.descripcion,
    dom.form.estilo,
    dom.form.tela,
    dom.form.realizadoPor,
    dom.form.rutasProcesos,
  ].forEach((input) => input.addEventListener("input", handleUppercaseTextInput));
  dom.form.produccionEstimada.addEventListener("focus", handleProductionEstimadaFocus);
  dom.form.produccionEstimada.addEventListener("input", handleProductionEstimadaInput);
  dom.form.produccionEstimada.addEventListener("blur", handleProductionEstimadaBlur);
  dom.searchProtoInput.addEventListener("input", handleSearchProtoInput);
  dom.searchProtoInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      handleSearch();
    }
  });
  dom.addRowBtn.addEventListener("click", () => appendRow());
  dom.removeRowBtn.addEventListener("click", removeLastRow);
  dom.clearTableBtn.addEventListener("click", clearTableWithConfirm);
  dom.addCorteRowBtn.addEventListener("click", () => appendCorteRow());
  dom.removeCorteRowBtn.addEventListener("click", removeLastCorteRow);
  dom.clearCorteTableBtn.addEventListener("click", clearCorteTableWithConfirm);
  dom.addAcabadoRowBtn.addEventListener("click", () => appendAcabadoRow());
  dom.removeAcabadoRowBtn.addEventListener("click", removeLastAcabadoRow);
  dom.clearAcabadoTableBtn.addEventListener("click", clearAcabadoTableWithConfirm);
  
  if (dom.tabCostura) dom.tabCostura.addEventListener("click", handleTabSwitch);
  if (dom.tabCorte) dom.tabCorte.addEventListener("click", handleTabSwitch);
  if (dom.tabAcabado) dom.tabAcabado.addEventListener("click", handleTabSwitch);
  if (dom.tabResumen) dom.tabResumen.addEventListener("click", handleTabSwitch);

  dom.masterSaveBtn.addEventListener("click", handleSaveMasterRow);
  dom.masterExcelBtn.addEventListener("click", handleExportMasterExcel);
  dom.masterAddBtn.addEventListener("click", handleAddMasterRow);
  dom.masterFilterTextInput.addEventListener("input", handleMasterFilterTextInput);
  dom.masterFilterTextInput.addEventListener("keydown", handleMasterFilterTextKeydown);
  dom.masterFilterSelect.addEventListener("change", handleMasterFilterSelectChange);
  dom.masterFilterClearBtn.addEventListener("click", handleMasterFilterClear);
  dom.saveBtn.addEventListener("click", handleSave);
  dom.copyPrototypeBtn.addEventListener("click", openCopyPrototypeModal);
  dom.clearPrototypeBtn.addEventListener("click", handleClearPrototypeWithConfirm);
  dom.confirmModalCancelBtn.addEventListener("click", closeConfirmModal);
  dom.confirmModalAcceptBtn.addEventListener("click", handleConfirmAccept);
  dom.copyPrototypeCancelBtn.addEventListener("click", closeCopyPrototypeModal);
  dom.copyPrototypeAcceptBtn.addEventListener("click", handleCopyPrototypeAccept);
  dom.copyPrototypeInput.addEventListener("input", handleProtoInput);
  dom.copyPrototypeInput.addEventListener("keydown", handleCopyPrototypeInputKeydown);
  dom.codeContextMenu.addEventListener("click", handleCodeContextMenuAction);
  document.addEventListener("click", handleDocumentClick);
  document.addEventListener("contextmenu", handleDocumentContextMenu);
  document.addEventListener("scroll", closeFloatingMenus, true);
  window.addEventListener("resize", closeFloatingMenus);
  dom.confirmModal.addEventListener("click", (event) => {
    if (event.target === dom.confirmModal) {
      closeConfirmModal();
    }
  });
  dom.copyPrototypeModal.addEventListener("click", (event) => {
    if (event.target === dom.copyPrototypeModal) {
      closeCopyPrototypeModal();
    }
  });
  document.addEventListener("keydown", (event) => {
    if (
      event.key === "Escape" &&
      ((dom.codeContextMenu && !dom.codeContextMenu.classList.contains("hidden")) ||
        (dom.masterFilterMenu && !dom.masterFilterMenu.classList.contains("hidden")))
    ) {
      closeFloatingMenus();
      return;
    }

    if (event.key === "Escape" && dom.confirmModal && !dom.confirmModal.classList.contains("hidden")) {
      closeConfirmModal();
      return;
    }

    if (event.key === "Escape" && dom.copyPrototypeModal && !dom.copyPrototypeModal.classList.contains("hidden")) {
      closeCopyPrototypeModal();
    }
  });
  window.addEventListener("afterprint", clearPrintRecordMode);
}

function normalizeSearchProto(value) {
  return String(value || "")
    .replace(/\D/g, "")
    .slice(0, 5);
}

function sanitizeIntegerInput(value, maxLength = Number.POSITIVE_INFINITY) {
  return String(value || "")
    .replace(/\D/g, "")
    .slice(0, maxLength);
}

function normalizeProtoValue(value) {
  return sanitizeIntegerInput(value, 5);
}

function normalizeShortIntegerValue(value) {
  return sanitizeIntegerInput(value, 2);
}

function normalizeIdemValue(value) {
  return sanitizeIntegerInput(value, 5);
}

function normalizeUppercaseTextValue(value) {
  return String(value || "").toLocaleUpperCase("es-PE");
}

function setUppercaseInputValue(input, value) {
  if (!input) {
    return;
  }

  input.value = normalizeUppercaseTextValue(value);
}

function normalizeProductionDigits(value) {
  return sanitizeIntegerInput(value);
}

function formatProductionEstimadaValue(value) {
  return AppUtils.formatInteger(normalizeProductionDigits(value));
}

function handleProtoInput(event) {
  const normalizedValue = normalizeProtoValue(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleVersionInput(event) {
  const normalizedValue = normalizeShortIntegerValue(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleIdemInput(event) {
  const normalizedValue = normalizeIdemValue(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleUppercaseTextInput(event) {
  const normalizedValue = normalizeUppercaseTextValue(event.target.value);

  if (event.target.value !== normalizedValue) {
    const selectionStart = event.target.selectionStart;
    const selectionEnd = event.target.selectionEnd;
    event.target.value = normalizedValue;

    if (
      typeof event.target.setSelectionRange === "function" &&
      selectionStart !== null &&
      selectionEnd !== null
    ) {
      event.target.setSelectionRange(selectionStart, selectionEnd);
    }
  }
}

function handleProductionEstimadaFocus(event) {
  if (event.target.readOnly || event.target.disabled) {
    return;
  }

  const normalizedValue = normalizeProductionDigits(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleProductionEstimadaInput(event) {
  if (event.target.readOnly || event.target.disabled) {
    return;
  }

  const normalizedValue = normalizeProductionDigits(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleProductionEstimadaBlur(event) {
  if (event.target.readOnly || event.target.disabled) {
    return;
  }

  event.target.value = formatProductionEstimadaValue(event.target.value);
}

function normalizeDecimalInputValue(value, decimals = 2) {
  const rawValue = String(value || "").replace(/[^\d.,]/g, "");
  const separators = rawValue.match(/[.,]/g) || [];

  if (!separators.length) {
    return rawValue.replace(/\D/g, "");
  }

  const separatorIndex = Math.max(rawValue.lastIndexOf("."), rawValue.lastIndexOf(","));
  const integerPart = rawValue.slice(0, separatorIndex).replace(/\D/g, "");
  const decimalDigits = rawValue.slice(separatorIndex + 1).replace(/\D/g, "");
  const normalizedInteger = integerPart || "0";

  if (separators.length === 1 && decimalDigits.length === 3) {
    return `${normalizedInteger}${decimalDigits}`;
  }

  if (separators.length > 1 && decimalDigits.length > decimals) {
    return rawValue.replace(/[^\d]/g, "");
  }

  const decimalPart = decimalDigits.slice(0, decimals);

  return decimalPart ? `${normalizedInteger}.${decimalPart}` : `${normalizedInteger}.`;
}

function normalizeProtectionInputValue(value) {
  return normalizeDecimalInputValue(value, 2);
}

function roundToDecimals(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round(AppUtils.safeNumber(value) * factor) / factor;
}

function normalizeProtectionValue(value) {
  const numericValue = roundToDecimals(AppUtils.safeNumber(value), 2);
  return numericValue > 0 ? numericValue : APP_CONFIG.defaultProtection;
}

function resolveActiveProtectionValue(rows, fallbackValue = APP_CONFIG.defaultProtection) {
  const match = (rows || []).find((row) => {
    const rawValue = row?.proteccion;
    return rawValue !== "" && rawValue !== null && rawValue !== undefined && AppUtils.safeNumber(rawValue) > 0;
  });
  return match ? normalizeProtectionValue(match.proteccion) : normalizeProtectionValue(fallbackValue);
}

function lockProtectionEditing(input) {
  if (!input) {
    return;
  }

  input.readOnly = true;
  input.tabIndex = -1;
  input.classList.remove("is-inline-editing");
  delete input.dataset.previousValue;
}

function enableProtectionEditing(input) {
  if (!input || !isEditableMode()) {
    return;
  }

  input.dataset.previousValue = input.value;
  input.readOnly = false;
  input.tabIndex = 0;
  input.classList.add("is-inline-editing");
  input.focus();
  input.select();
}

function updateRowProtection(row, protectionValue, options = {}) {
  const rowData = readRowValues(row);
  const normalizedProtection = normalizeProtectionValue(protectionValue);
  const isManual = rowData.tipoPta === "*";
  const tiempoMaq = isManual ? 0 : rowData.tiempoEstimado * normalizedProtection;
  const tiempoManual = isManual ? rowData.tiempoEstimado * normalizedProtection : 0;

  writeRowValues(row, {
    ...rowData,
    proteccion: normalizedProtection,
    tiempoMaq,
    tiempoManual,
    tiempoCotizacion: tiempoMaq + tiempoManual,
  });

  if (options.refreshSummary !== false) {
    refreshSummary();
  }
}

function updateCorteProtection(row, protectionValue, options = {}) {
  const data = readCorteRowValues(row);
  const normalizedProtection = normalizeProtectionValue(protectionValue);
  const tcExt = AppUtils.evaluateFormula(data.tiempoEstimadoCorte);
  const thExt = AppUtils.evaluateFormula(data.tiempoEstimadoHabilitado);
  const isCorte = data.area.toUpperCase() === "CORT";
  const isHab = data.area.toUpperCase() === "HAB";
  const hasCorteEstimate = Boolean(data.tiempoEstimadoCorte);
  const hasHabEstimate = Boolean(data.tiempoEstimadoHabilitado);
  const tiempoCorte = isCorte && hasCorteEstimate ? tcExt * normalizedProtection : "";
  const tiempoHab = isHab && hasHabEstimate ? thExt * normalizedProtection : "";
  const tiempoCotizacion =
    tiempoCorte === "" && tiempoHab === ""
      ? ""
      : AppUtils.safeNumber(tiempoCorte) + AppUtils.safeNumber(tiempoHab);

  writeCorteRowValues(row, {
    ...data,
    proteccion: normalizedProtection,
    tiempoCorte,
    tiempoHab,
    tiempoCotizacion,
  });

  if (options.refreshSummary !== false) {
    refreshCorteSummary();
  }
}

function updateAcabadoProtection(row, protectionValue, options = {}) {
  const data = readAcabadoRowValues(row);
  const normalizedProtection = normalizeProtectionValue(protectionValue);
  const tiempoCotizacion =
    data.tiempoEstimado === "" ? "" : data.tiempoEstimado * normalizedProtection;

  writeAcabadoRowValues(row, {
    ...data,
    proteccion: normalizedProtection,
    tiempoCotizacion,
  });

  if (options.refreshSummary !== false) {
    refreshAcabadoSummary();
  }
}

function getProtectionTableContext(row) {
  if (!row) {
    return null;
  }

  if (row.querySelector('[data-field="codigo"]')) {
    return {
      isEditableTarget: true,
      protectionStateKey: "costuraProtectionValue",
      refreshSummary,
      rows: getTableRows(),
      updateRow: updateRowProtection,
    };
  }

  if (row.querySelector('[data-field="tiempoEstimadoCorte"]')) {
    return {
      isEditableTarget: !isCorteRowEmpty(row),
      protectionStateKey: "corteProtectionValue",
      refreshSummary: refreshCorteSummary,
      rows: Array.from(dom.corteTableBody?.querySelectorAll("tr") || []),
      updateRow: updateCorteProtection,
    };
  }

  if (row.querySelector('[data-field="tiempoEstimado"]')) {
    return {
      isEditableTarget: !isAcabadoRowEmpty(row),
      protectionStateKey: "acabadoProtectionValue",
      refreshSummary: refreshAcabadoSummary,
      rows: Array.from(dom.acabadoTableBody?.querySelectorAll("tr") || []),
      updateRow: updateAcabadoProtection,
    };
  }

  return null;
}

function commitProtectionEdit(input, options = {}) {
  if (!input) {
    return;
  }

  const row = input.closest("tr");
  const tableContext = getProtectionTableContext(row);
  const previousValue = AppUtils.safeNumber(input.dataset.previousValue) || APP_CONFIG.defaultProtection;
  const rawValue = String(input.value || "").trim();
  const normalizedValue = normalizeProtectionInputValue(rawValue);
  const hasDigits = /\d/.test(normalizedValue);
  let nextProtection = previousValue;

  if (!options.restorePrevious) {
    if (hasDigits) {
      nextProtection = normalizeProtectionValue(normalizedValue);
    }

    if (!hasDigits || nextProtection <= 0) {
      nextProtection = previousValue;

      if (rawValue) {
        showToast("Ingresa un valor válido para % Protección con hasta 2 decimales.", "error");
      }
    }
  }

  lockProtectionEditing(input);

  if (!tableContext) {
    return;
  }

  if (tableContext.protectionStateKey) {
    appState[tableContext.protectionStateKey] = normalizeProtectionValue(nextProtection);
  }

  tableContext.rows.forEach((targetRow) => {
    tableContext.updateRow(targetRow, nextProtection, { refreshSummary: false });
  });
  tableContext.refreshSummary();
}

function handleProtectionDoubleClick(event) {
  if (!isEditableMode()) {
    return;
  }

  const row = event.currentTarget.closest("tr");
  const tableContext = getProtectionTableContext(row);

  if (!tableContext || !tableContext.isEditableTarget) {
    return;
  }

  enableProtectionEditing(event.currentTarget);
}

function handleProtectionInput(event) {
  const normalizedValue = normalizeProtectionInputValue(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function handleProtectionBlur(event) {
  if (!event.currentTarget.classList.contains("is-inline-editing")) {
    return;
  }

  commitProtectionEdit(event.currentTarget);
}

function handleProtectionKeydown(event) {
  if (event.key === "Enter") {
    event.preventDefault();
    event.currentTarget.blur();
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    const previousValue = event.currentTarget.dataset.previousValue || "";
    event.currentTarget.value = previousValue;
    event.currentTarget.blur();
  }
}

function normalizeGeneralFormFields() {
  dom.form.proto.value = normalizeProtoValue(dom.form.proto.value);
  dom.form.version.value = normalizeShortIntegerValue(dom.form.version.value);
  dom.form.idem.value = normalizeIdemValue(dom.form.idem.value);
  setUppercaseInputValue(dom.form.clienteCustom, dom.form.clienteCustom.value);
  setUppercaseInputValue(dom.form.descripcion, dom.form.descripcion.value);
  setUppercaseInputValue(dom.form.estilo, dom.form.estilo.value);
  setUppercaseInputValue(dom.form.tela, dom.form.tela.value);
  setUppercaseInputValue(dom.form.realizadoPor, dom.form.realizadoPor.value);
  setUppercaseInputValue(dom.form.rutasProcesos, dom.form.rutasProcesos.value);
  dom.form.produccionEstimada.value = formatProductionEstimadaValue(dom.form.produccionEstimada.value);
}

function handleDocumentClick(event) {
  const isCodeMenuOpen = dom.codeContextMenu && !dom.codeContextMenu.classList.contains("hidden");
  const isMasterFilterOpen = dom.masterFilterMenu && !dom.masterFilterMenu.classList.contains("hidden");

  if (!isCodeMenuOpen && !isMasterFilterOpen) {
    return;
  }

  if (isCodeMenuOpen && dom.codeContextMenu.contains(event.target)) {
    return;
  }

  if (isMasterFilterOpen && dom.masterFilterMenu.contains(event.target)) {
    return;
  }

  closeFloatingMenus();
}

function handleDocumentContextMenu(event) {
  if (event.target.closest(".code-input") || event.target.closest(".master-filter-menu")) {
    return;
  }

  closeFloatingMenus();
}

function positionFloatingMenu(menu, x, y) {
  if (!menu) {
    return;
  }

  const margin = 12;
  const menuWidth = menu.offsetWidth;
  const menuHeight = menu.offsetHeight;
  const left = Math.min(x, window.innerWidth - menuWidth - margin);
  const top = Math.min(y, window.innerHeight - menuHeight - margin);

  menu.style.left = `${Math.max(margin, left)}px`;
  menu.style.top = `${Math.max(margin, top)}px`;
  menu.style.visibility = "visible";
}

function closeFloatingMenus() {
  closeCodeContextMenu();
  closeMasterFilterMenu();
}

function openCodeContextMenu(row, x, y, tableType = "costura") {
  if (!dom.codeContextMenu || !row) {
    return;
  }

  closeMasterFilterMenu();
  codeContextMenuRow = row;
  codeContextMenuTable = tableType || "costura";
  dom.codeContextMenu.classList.remove("hidden");
  dom.codeContextMenu.setAttribute("aria-hidden", "false");
  dom.codeContextMenu.style.visibility = "hidden";
  dom.codeContextMenu.style.left = "0px";
  dom.codeContextMenu.style.top = "0px";
  positionFloatingMenu(dom.codeContextMenu, x, y);
}

function closeCodeContextMenu() {
  if (!dom.codeContextMenu) {
    return;
  }

  codeContextMenuRow = null;
  dom.codeContextMenu.classList.add("hidden");
  dom.codeContextMenu.setAttribute("aria-hidden", "true");
  dom.codeContextMenu.style.visibility = "";
  dom.codeContextMenu.style.left = "";
  dom.codeContextMenu.style.top = "";
}

function closeMasterFilterMenu() {
  if (!dom.masterFilterMenu) {
    return;
  }

  dom.masterFilterMenu.dataset.comparableHeaderKey = "";
  dom.masterFilterMenu.dataset.filterType = "";
  dom.masterFilterMenu.classList.add("hidden");
  dom.masterFilterMenu.setAttribute("aria-hidden", "true");
  dom.masterFilterMenu.style.visibility = "";
  dom.masterFilterMenu.style.left = "";
  dom.masterFilterMenu.style.top = "";
}

function getMasterFilterType(comparableHeaderKey) {
  if (MASTER_TEXT_FILTER_HEADERS.has(comparableHeaderKey)) {
    return "text";
  }

  if (MASTER_SELECT_FILTER_HEADERS.has(comparableHeaderKey)) {
    return "select";
  }

  return "";
}

function getMasterFilterValue(comparableHeaderKey) {
  return String(appState.masterFilters[comparableHeaderKey] || "").trim();
}

function getMasterActiveFiltersCount() {
  return Object.values(appState.masterFilters).filter((value) => String(value || "").trim()).length;
}

function getMasterHeaderByComparableKey(comparableHeaderKey) {
  return (
    appState.masterHeaders.find((header) => getMasterComparableKey(header) === comparableHeaderKey) || ""
  );
}

function getMasterLoadedSummaryMessage() {
  const totalRows = appState.masterSourceRows.length;
  const visibleRows = appState.masterRows.length;
  const activeFilters = getMasterActiveFiltersCount();
  const countLabel = activeFilters
    ? `${visibleRows} de ${totalRows} operaciones visibles.`
    : `${totalRows} operaciones.`;
  const filterLabel = activeFilters ? ` ${activeFilters} filtro(s) activo(s).` : "";

  return `${countLabel}${filterLabel} Click derecho en descripcion abreviada para editar una fila.`;
}

function updateMasterLoadedSummary() {
  updateMasterInfo(getMasterLoadedSummaryMessage());
}

function getMasterComparableHeaderKeyFromFilterMenu() {
  return String(dom.masterFilterMenu?.dataset.comparableHeaderKey || "");
}

function getMasterUniqueFilterValues(comparableHeaderKey) {
  const header = getMasterHeaderByComparableKey(comparableHeaderKey);

  if (!header) {
    return [];
  }

  const headerKey = AppUtils.normalizeKey(header);
  const uniqueValues = new Map();

  appState.masterSourceRows.forEach((row) => {
    const rawValue = String(row.values[headerKey] || "").trim();

    if (!rawValue) {
      return;
    }

    const normalizedValue = AppUtils.normalizeKey(rawValue);

    if (!uniqueValues.has(normalizedValue)) {
      uniqueValues.set(normalizedValue, rawValue);
    }
  });

  return Array.from(uniqueValues.values()).sort((left, right) => VERSION_COLLATOR.compare(left, right));
}

function getMasterCodeHeaderKey() {
  const header = getMasterHeaderByComparableKey("CODIGO");
  return header ? AppUtils.normalizeKey(header) : "";
}

function sortMasterRows(rows = []) {
  const codeHeaderKey = getMasterCodeHeaderKey();

  if (!codeHeaderKey) {
    return [...rows];
  }

  return [...rows].sort((left, right) => {
    const leftCode = String(left?.values?.[codeHeaderKey] || "").trim();
    const rightCode = String(right?.values?.[codeHeaderKey] || "").trim();

    if (leftCode && rightCode) {
      const codeCompare = VERSION_COLLATOR.compare(leftCode, rightCode);
      if (codeCompare !== 0) {
        return codeCompare;
      }
    } else if (leftCode) {
      return -1;
    } else if (rightCode) {
      return 1;
    }

    const leftRowNumber = Number(left?.rowNumber) || 0;
    const rightRowNumber = Number(right?.rowNumber) || 0;

    if (leftRowNumber !== rightRowNumber) {
      return leftRowNumber - rightRowNumber;
    }

    return VERSION_COLLATOR.compare(String(left?.key || ""), String(right?.key || ""));
  });
}

function applyMasterFilters() {
  const activeFilters = Object.entries(appState.masterFilters).filter(([, value]) => String(value || "").trim());

  if (!activeFilters.length) {
    appState.masterRows = sortMasterRows(appState.masterSourceRows);
    return;
  }

  appState.masterRows = sortMasterRows(appState.masterSourceRows.filter((row) => {
    if (row.key === appState.masterEditingRowKey) {
      return true;
    }

    return activeFilters.every(([comparableHeaderKey, filterValue]) => {
      const header = getMasterHeaderByComparableKey(comparableHeaderKey);

      if (!header) {
        return true;
      }

      const headerKey = AppUtils.normalizeKey(header);
      const cellValue = String(row.values[headerKey] || "").trim();

      if (MASTER_TEXT_FILTER_HEADERS.has(comparableHeaderKey)) {
        return AppUtils.normalizeKey(cellValue).includes(AppUtils.normalizeKey(filterValue));
      }

      return AppUtils.normalizeKey(cellValue) === AppUtils.normalizeKey(filterValue);
    });
  }));
}

function setMasterFilterValue(comparableHeaderKey, value) {
  const normalizedValue = String(value || "").trim();

  if (normalizedValue) {
    appState.masterFilters[comparableHeaderKey] = normalizedValue;
  } else {
    delete appState.masterFilters[comparableHeaderKey];
  }

  applyMasterFilters();
  renderMasterTable();
  updateMasterLoadedSummary();
}

function syncMasterFilterMenuState() {
  if (!dom.masterFilterMenu) {
    return;
  }

  const comparableHeaderKey = getMasterComparableHeaderKeyFromFilterMenu();
  const filterType = String(dom.masterFilterMenu.dataset.filterType || "");
  const filterValue = getMasterFilterValue(comparableHeaderKey);

  dom.masterFilterClearBtn.disabled = !filterValue;

  if (filterType === "text") {
    dom.masterFilterTextInput.value = filterValue;
  }

  if (filterType === "select") {
    dom.masterFilterSelect.value = filterValue;
  }
}

function openMasterFilterMenu(header, x, y) {
  if (!dom.masterFilterMenu) {
    return;
  }

  const comparableHeaderKey = getMasterComparableKey(header);
  const filterType = getMasterFilterType(comparableHeaderKey);

  if (!filterType) {
    return;
  }

  closeCodeContextMenu();
  dom.masterFilterMenu.dataset.comparableHeaderKey = comparableHeaderKey;
  dom.masterFilterMenu.dataset.filterType = filterType;
  dom.masterFilterMenu.classList.remove("hidden");
  dom.masterFilterMenu.setAttribute("aria-hidden", "false");
  dom.masterFilterMenu.style.visibility = "hidden";
  dom.masterFilterMenu.style.left = "0px";
  dom.masterFilterMenu.style.top = "0px";

  dom.masterFilterTitle.textContent = `Filtrar ${formatMasterHeaderLabel(header)}`;
  dom.masterFilterLabel.textContent =
    filterType === "text" ? "Escribe el valor a buscar" : "Selecciona un valor de la lista";
  dom.masterFilterTextInput.classList.toggle("hidden", filterType !== "text");
  dom.masterFilterSelect.classList.toggle("hidden", filterType !== "select");

  if (filterType === "select") {
    const currentFilterValue = getMasterFilterValue(comparableHeaderKey);
    const allOption = document.createElement("option");
    allOption.value = "";
    allOption.textContent = "Todos";

    dom.masterFilterSelect.innerHTML = "";
    dom.masterFilterSelect.appendChild(allOption);
    getMasterUniqueFilterValues(comparableHeaderKey).forEach((optionValue) => {
      const option = document.createElement("option");
      option.value = optionValue;
      option.textContent = optionValue;
      dom.masterFilterSelect.appendChild(option);
    });
    dom.masterFilterSelect.value = currentFilterValue;
  }

  syncMasterFilterMenuState();
  positionFloatingMenu(dom.masterFilterMenu, x, y);

  if (filterType === "text") {
    dom.masterFilterTextInput.focus();
    dom.masterFilterTextInput.select();
    return;
  }

  dom.masterFilterSelect.focus();
}

function handleCodeContextMenu(event) {
  if (!isEditableMode()) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  openCodeContextMenu(event.currentTarget.closest("tr"), event.clientX, event.clientY);
}

function handleOperationsContextMenu(event, tableType) {
  if (!isEditableMode()) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  openCodeContextMenu(event.currentTarget.closest("tr"), event.clientX, event.clientY, tableType);
}

function handleCodeContextMenuAction(event) {
  const actionButton = event.target.closest("[data-action]");

  if (!actionButton || !codeContextMenuRow) {
    return;
  }

  const targetRow = codeContextMenuRow;
  const tableType = codeContextMenuTable;
  closeCodeContextMenu();

  if (tableType === "corte") {
    if (actionButton.dataset.action === "insert-above") {
      insertCorteRowAbove(targetRow);
    } else if (actionButton.dataset.action === "insert-below") {
      insertCorteRowBelow(targetRow);
    } else if (actionButton.dataset.action === "delete-row") {
      deleteCorteTableRow(targetRow);
    }
    return;
  }

  if (tableType === "acabado") {
    if (actionButton.dataset.action === "insert-above") {
      insertAcabadoRowAbove(targetRow);
    } else if (actionButton.dataset.action === "insert-below") {
      insertAcabadoRowBelow(targetRow);
    } else if (actionButton.dataset.action === "delete-row") {
      deleteAcabadoTableRow(targetRow);
    }
    return;
  }

  if (actionButton.dataset.action === "insert-above") {
    insertRowAbove(targetRow);
    return;
  }

  if (actionButton.dataset.action === "insert-below") {
    insertRowBelow(targetRow);
    return;
  }

  if (actionButton.dataset.action === "delete-row") {
    deleteTableRow(targetRow);
  }
}

function handleMasterHeaderContextMenu(event) {
  if (appState.currentView !== "master") {
    return;
  }

  const header = event.currentTarget.dataset.header || "";
  const comparableHeaderKey = event.currentTarget.dataset.comparableHeaderKey || "";

  if (!getMasterFilterType(comparableHeaderKey)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  if (appState.masterEditingRowKey) {
    showToast("Guarda la fila actual antes de usar los filtros del Maestro de operaciones.", "info");
    focusMasterEditingRow();
    return;
  }

  openMasterFilterMenu(header, event.clientX, event.clientY);
}

function handleMasterFilterTextInput(event) {
  const comparableHeaderKey = getMasterComparableHeaderKeyFromFilterMenu();

  if (!comparableHeaderKey) {
    return;
  }

  setMasterFilterValue(comparableHeaderKey, event.currentTarget.value);
  syncMasterFilterMenuState();
}

function handleMasterFilterTextKeydown(event) {
  if (event.key === "Enter") {
    event.preventDefault();
    closeMasterFilterMenu();
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    closeMasterFilterMenu();
  }
}

function handleMasterFilterSelectChange(event) {
  const comparableHeaderKey = getMasterComparableHeaderKeyFromFilterMenu();

  if (!comparableHeaderKey) {
    return;
  }

  setMasterFilterValue(comparableHeaderKey, event.currentTarget.value);
  syncMasterFilterMenuState();
}

function handleMasterFilterClear() {
  const comparableHeaderKey = getMasterComparableHeaderKeyFromFilterMenu();

  if (!comparableHeaderKey) {
    return;
  }

  setMasterFilterValue(comparableHeaderKey, "");
  syncMasterFilterMenuState();

  if (dom.masterFilterMenu.dataset.filterType === "text") {
    dom.masterFilterTextInput.focus();
    return;
  }

  dom.masterFilterSelect.focus();
}

function formatMasterHeaderLabel(header) {
  return String(header || "")
    .replace(/_/g, " ")
    .trim();
}

function getMasterComparableKey(value) {
  return AppUtils.normalizeKey(value).replace(/\s+/g, "_");
}

function getMasterColumnWeight(headerKey) {
  const comparableKey = getMasterComparableKey(headerKey);
  const weightMap = {
    CODIGO: 0.42,
    TIEMPO: 0.42,
    BLOQUE: 0.5,
    TIPO_OPERACION: 0.58,
    ACCION: 0.48,
    TIPO_MAQ: 0.48,
    ESTADO: 0.42,
    DESCRIPCION_ABREVIADA: 1.9,
    DESCRIPCION_COMPLETA: 1.95,
    SECCION: 0.68,
    MAQUINA: 1.05,
    FECHA_REGISTRO: 0.54,
  };

  return weightMap[comparableKey] || 0.72;
}

function renderMasterColgroup() {
  if (!dom.masterTableColgroup) {
    return;
  }

  dom.masterTableColgroup.innerHTML = "";

  if (!appState.masterHeaders.length) {
    return;
  }

  const columnDefinitions = appState.masterHeaders.map((header) => ({
    headerKey: AppUtils.normalizeKey(header),
    weight: getMasterColumnWeight(header),
  }));
  const totalWeight = columnDefinitions.reduce((sum, item) => sum + item.weight, 0);

  columnDefinitions.forEach((column) => {
    const col = document.createElement("col");
    col.style.width = `${((column.weight / totalWeight) * 100).toFixed(4)}%`;
    dom.masterTableColgroup.appendChild(col);
  });
}

function getMasterRowByKey(rowKey) {
  return appState.masterSourceRows.find((row) => row.key === rowKey) || null;
}

function padMasterDatePart(value) {
  return String(value).padStart(2, "0");
}

function parseMasterDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getTime());
  }

  const rawValue = String(value || "").trim();

  if (!rawValue) {
    return null;
  }

  const partsMatch = rawValue.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );

  if (partsMatch) {
    const day = Number(partsMatch[1]);
    const month = Number(partsMatch[2]);
    const year = Number(partsMatch[3].length === 2 ? `20${partsMatch[3]}` : partsMatch[3]);
    const hours = Number(partsMatch[4] || 0);
    const minutes = Number(partsMatch[5] || 0);
    const seconds = Number(partsMatch[6] || 0);
    const parsedDate = new Date(year, month - 1, day, hours, minutes, seconds);

    if (!Number.isNaN(parsedDate.getTime())) {
      return parsedDate;
    }
  }

  const fallbackDate = new Date(rawValue);
  return Number.isNaN(fallbackDate.getTime()) ? null : fallbackDate;
}

function formatMasterDateTimeValue(dateValue) {
  const parsedDate = parseMasterDateValue(dateValue) || new Date();

  return [
    `${padMasterDatePart(parsedDate.getDate())}/${padMasterDatePart(parsedDate.getMonth() + 1)}/${parsedDate.getFullYear()}`,
    `${padMasterDatePart(parsedDate.getHours())}:${padMasterDatePart(parsedDate.getMinutes())}:${padMasterDatePart(parsedDate.getSeconds())}`,
  ].join(" ");
}

function formatMasterDateShortValue(value) {
  const parsedDate = parseMasterDateValue(value);

  if (!parsedDate) {
    return String(value || "");
  }

  const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const shortYear = String(parsedDate.getFullYear()).slice(-2);

  return `${padMasterDatePart(parsedDate.getDate())}/${monthNames[parsedDate.getMonth()]}/${shortYear}`;
}

function formatMasterDateTooltip(value) {
  const parsedDate = parseMasterDateValue(value);

  if (!parsedDate) {
    return "";
  }

  const hours = parsedDate.getHours();
  const minutes = padMasterDatePart(parsedDate.getMinutes());
  const meridiem = hours >= 12 ? "pm" : "am";
  const twelveHour = hours % 12 || 12;

  return `Hora: ${padMasterDatePart(twelveHour)}:${minutes}${meridiem}`;
}

function isMasterFullContentTooltipColumn(comparableHeaderKey) {
  return (
    comparableHeaderKey === "DESCRIPCION_ABREVIADA" ||
    comparableHeaderKey === "DESCRIPCION_COMPLETA" ||
    comparableHeaderKey === "MAQUINA"
  );
}

function isMasterCenteredColumn(comparableHeaderKey) {
  return (
    comparableHeaderKey === "CODIGO" ||
    comparableHeaderKey === "TIEMPO" ||
    comparableHeaderKey === "SECCION" ||
    comparableHeaderKey === "BLOQUE" ||
    comparableHeaderKey === "TIPO_OPERACION" ||
    comparableHeaderKey === "ACCION" ||
    comparableHeaderKey === "TIPO_MAQ" ||
    comparableHeaderKey === "MAQUINA" ||
    comparableHeaderKey === "ESTADO" ||
    comparableHeaderKey === MASTER_DATE_HEADER
  );
}

function getMasterCurrentDatePreview() {
  return formatMasterDateTimeValue(new Date());
}

function buildMasterRowKey(rowData = {}) {
  if (rowData.rowNumber) {
    return `master-row-${rowData.rowNumber}`;
  }

  masterDraftSequence += 1;
  return `master-draft-${masterDraftSequence}`;
}

function updateMasterInfo(message = "") {
  if (!dom.masterInfo) {
    return;
  }

  dom.masterInfo.textContent = message;
}

function setMasterLoadingState(isLoading) {
  appState.masterIsLoading = Boolean(isLoading);
  renderMasterTable();
}

function updateMasterActionState() {
  if (!dom.masterSaveBtn || !dom.masterExcelBtn || !dom.masterAddBtn) {
    return;
  }

  const hasHeaders = appState.masterHeaders.length > 0;
  const hasEditingRow = Boolean(appState.masterEditingRowKey);

  dom.masterSaveBtn.disabled = !hasEditingRow;
  dom.masterExcelBtn.disabled = !hasHeaders || appState.masterIsLoading;
  dom.masterAddBtn.disabled = !hasHeaders || hasEditingRow;
}

function getMasterExcelFileName() {
  const now = new Date();
  const year = now.getFullYear();
  const month = padMasterDatePart(now.getMonth() + 1);
  const day = padMasterDatePart(now.getDate());
  const hours = padMasterDatePart(now.getHours());
  const minutes = padMasterDatePart(now.getMinutes());
  const seconds = padMasterDatePart(now.getSeconds());

  return `maestro_operaciones_${year}${month}${day}_${hours}${minutes}${seconds}.xls`;
}

function inlineElementStyles(sourceElement, targetElement) {
  if (!sourceElement || !targetElement) {
    return;
  }

  const computedStyles = window.getComputedStyle(sourceElement);

  Array.from(computedStyles).forEach((propertyName) => {
    targetElement.style.setProperty(
      propertyName,
      computedStyles.getPropertyValue(propertyName),
      computedStyles.getPropertyPriority(propertyName)
    );
  });
}

function getControlDisplayValue(control) {
  if (!control) {
    return "";
  }

  if (control.tagName === "SELECT") {
    const selectedOption = control.options[control.selectedIndex];
    return selectedOption ? selectedOption.textContent.trim() : "";
  }

  return String(control.value ?? "").trim();
}

function createMasterExportValueNode(sourceControl) {
  const valueNode = document.createElement("div");

  inlineElementStyles(sourceControl, valueNode);
  valueNode.textContent = getControlDisplayValue(sourceControl);
  valueNode.style.display = "block";
  valueNode.style.width = "100%";
  valueNode.style.border = "none";
  valueNode.style.background = "transparent";
  valueNode.style.boxShadow = "none";
  valueNode.style.overflow = "hidden";
  valueNode.style.textOverflow = "ellipsis";
  valueNode.style.whiteSpace = "nowrap";
  valueNode.style.minHeight = `${Math.max(sourceControl.offsetHeight || 0, 34)}px`;

  return valueNode;
}

function buildMasterExcelTableClone() {
  const sourceTable = dom.masterPanel?.querySelector(".master-table");

  if (!sourceTable) {
    return null;
  }

  const clonedTable = sourceTable.cloneNode(true);
  const sourceElements = [sourceTable, ...sourceTable.querySelectorAll("*")];
  const clonedElements = [clonedTable, ...clonedTable.querySelectorAll("*")];

  sourceElements.forEach((sourceElement, index) => {
    inlineElementStyles(sourceElement, clonedElements[index]);
  });

  const sourceControls = sourceTable.querySelectorAll("input, select, textarea");
  const clonedControls = clonedTable.querySelectorAll("input, select, textarea");

  sourceControls.forEach((sourceControl, index) => {
    const clonedControl = clonedControls[index];

    if (!clonedControl) {
      return;
    }

    clonedControl.replaceWith(createMasterExportValueNode(sourceControl));
  });

  clonedTable.style.width = `${sourceTable.offsetWidth || sourceTable.scrollWidth || 0}px`;
  clonedTable.style.maxWidth = "none";
  clonedTable.style.tableLayout = "fixed";

  return clonedTable;
}

function buildMasterExcelDocument(tableMarkup) {
  return [
    "<!DOCTYPE html>",
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">',
    "<head>",
    '<meta charset="UTF-8">',
    '<meta http-equiv="X-UA-Compatible" content="IE=edge">',
    "<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Maestro</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->",
    "<style>",
    "body { margin: 18px; background: #ffffff; }",
    "table { border-collapse: collapse; }",
    "</style>",
    "</head>",
    "<body>",
    tableMarkup,
    "</body>",
    "</html>",
  ].join("");
}

function downloadMasterExcelFile(documentHtml) {
  const blob = new Blob(["\ufeff", documentHtml], {
    type: "application/vnd.ms-excel;charset=utf-8;",
  });
  const fileName = getMasterExcelFileName();

  if (window.navigator?.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, fileName);
    return;
  }

  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = downloadUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();

  window.setTimeout(() => {
    URL.revokeObjectURL(downloadUrl);
  }, 1200);
}

function handleExportMasterExcel() {
  if (!appState.masterHeaders.length || appState.masterIsLoading) {
    showToast("Primero carga el Maestro de operaciones antes de exportar.", "info");
    return;
  }

  const exportedTable = buildMasterExcelTableClone();

  if (!exportedTable) {
    showToast("No se pudo preparar la tabla del Maestro para exportar.", "error");
    return;
  }

  downloadMasterExcelFile(buildMasterExcelDocument(exportedTable.outerHTML));
  showToast("Se descargó el archivo Excel del Maestro de operaciones.", "success");
}

function getMasterExcelFileName() {
  const now = new Date();
  const year = now.getFullYear();
  const month = padMasterDatePart(now.getMonth() + 1);
  const day = padMasterDatePart(now.getDate());
  const hours = padMasterDatePart(now.getHours());
  const minutes = padMasterDatePart(now.getMinutes());
  const seconds = padMasterDatePart(now.getSeconds());

  return `maestro_operaciones_${year}${month}${day}_${hours}${minutes}${seconds}.xlsx`;
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeXmlText(value) {
  return escapeXml(String(value ?? "").replace(/\r\n?/g, "\n")).replace(/\n/g, "&#10;");
}

function getPrimaryFontName(fontFamily) {
  return (
    String(fontFamily || "")
      .split(",")[0]
      .replace(/['"]/g, "")
      .trim() || "Calibri"
  );
}

function pixelsToPoints(value) {
  return Math.round(((Number(value) || 0) * 72) / 96 * 100) / 100;
}

function pixelsToExcelWidth(value) {
  const pixels = Number(value) || 0;
  const width = pixels > 12 ? (pixels - 5) / 7 : pixels / 12;
  return Math.max(Math.round(width * 100) / 100, 0.1);
}

function cssColorToArgb(value) {
  const rawValue = String(value || "").trim();

  if (!rawValue || rawValue === "transparent") {
    return null;
  }

  const hexMatch = rawValue.match(/^#([0-9a-f]{3,8})$/i);

  if (hexMatch) {
    const hexValue = hexMatch[1];

    if (hexValue.length === 3) {
      const [r, g, b] = hexValue.split("");
      return `FF${r}${r}${g}${g}${b}${b}`.toUpperCase();
    }

    if (hexValue.length === 6) {
      return `FF${hexValue}`.toUpperCase();
    }

    if (hexValue.length === 8) {
      return hexValue.toUpperCase();
    }
  }

  const rgbMatch = rawValue.match(/^rgba?\(([^)]+)\)$/i);

  if (!rgbMatch) {
    return null;
  }

  const parts = rgbMatch[1].split(",").map((part) => part.trim());
  const red = Math.max(0, Math.min(255, Math.round(Number(parts[0]) || 0)));
  const green = Math.max(0, Math.min(255, Math.round(Number(parts[1]) || 0)));
  const blue = Math.max(0, Math.min(255, Math.round(Number(parts[2]) || 0)));
  const alphaValue = parts[3] === undefined ? 1 : Number(parts[3]);
  const alpha = Math.max(0, Math.min(255, Math.round((Number.isFinite(alphaValue) ? alphaValue : 1) * 255)));
  const toHex = (number) => number.toString(16).padStart(2, "0").toUpperCase();

  if (alpha === 0) {
    return null;
  }

  return `${toHex(alpha)}${toHex(red)}${toHex(green)}${toHex(blue)}`;
}

function resolveEffectiveBackgroundColor(elements, fallback = "FFFFFFFF") {
  for (const element of elements) {
    if (!element) {
      continue;
    }

    const argb = cssColorToArgb(window.getComputedStyle(element).backgroundColor);

    if (argb && !argb.startsWith("00")) {
      return argb;
    }
  }

  return fallback;
}

function resolveBorderSide(style, sideName) {
  const width = parseFloat(style.getPropertyValue(`border-${sideName}-width`) || "0");
  const borderStyle = style.getPropertyValue(`border-${sideName}-style`) || "none";
  const color = cssColorToArgb(style.getPropertyValue(`border-${sideName}-color`));

  if (width <= 0 || borderStyle === "none" || borderStyle === "hidden") {
    return null;
  }

  return {
    color: color || "FF000000",
    style: "thin",
  };
}

function mapHorizontalAlignment(value) {
  const normalizedValue = String(value || "").trim().toLowerCase();

  if (normalizedValue === "center") {
    return "center";
  }

  if (normalizedValue === "right" || normalizedValue === "end") {
    return "right";
  }

  if (normalizedValue === "justify") {
    return "justify";
  }

  return "left";
}

function getMasterCellDisplayText(cellElement) {
  const control = cellElement.querySelector("input, select, textarea");

  if (control) {
    return getControlDisplayValue(control);
  }

  const headerLines = cellElement.querySelectorAll(".master-header-line");

  if (headerLines.length) {
    return Array.from(headerLines)
      .map((line) => line.textContent.trim())
      .join("\n");
  }

  return String(cellElement.innerText ?? cellElement.textContent ?? "")
    .replace(/\r\n?/g, "\n")
    .trim();
}

function createMasterXlsxStyleRegistry() {
  return {
    fonts: [
      {
        bold: false,
        color: "FF000000",
        fontName: "Calibri",
        fontSize: 11,
        italic: false,
      },
    ],
    fills: [
      { patternType: "none" },
      { patternType: "gray125" },
    ],
    borders: [
      { bottom: null, left: null, right: null, top: null },
    ],
    cellXfs: [
      {
        alignment: null,
        borderId: 0,
        fillId: 0,
        fontId: 0,
      },
    ],
    fontMap: new Map(),
    fillMap: new Map([
      [JSON.stringify({ patternType: "none" }), 0],
      [JSON.stringify({ patternType: "gray125" }), 1],
    ]),
    borderMap: new Map([
      [JSON.stringify({ bottom: null, left: null, right: null, top: null }), 0],
    ]),
    xfMap: new Map([
      [JSON.stringify({ alignment: null, borderId: 0, fillId: 0, fontId: 0 }), 0],
    ]),
  };
}

function registerMasterXlsxFont(registry, fontDefinition) {
  const key = JSON.stringify(fontDefinition);

  if (!registry.fontMap.has(key)) {
    registry.fontMap.set(key, registry.fonts.length);
    registry.fonts.push(fontDefinition);
  }

  return registry.fontMap.get(key);
}

function registerMasterXlsxFill(registry, fillDefinition) {
  const key = JSON.stringify(fillDefinition);

  if (!registry.fillMap.has(key)) {
    registry.fillMap.set(key, registry.fills.length);
    registry.fills.push(fillDefinition);
  }

  return registry.fillMap.get(key);
}

function registerMasterXlsxBorder(registry, borderDefinition) {
  const key = JSON.stringify(borderDefinition);

  if (!registry.borderMap.has(key)) {
    registry.borderMap.set(key, registry.borders.length);
    registry.borders.push(borderDefinition);
  }

  return registry.borderMap.get(key);
}

function registerMasterXlsxCellStyle(registry, styleDefinition) {
  const fontId = registerMasterXlsxFont(registry, styleDefinition.font);
  const fillId = registerMasterXlsxFill(registry, styleDefinition.fill);
  const borderId = registerMasterXlsxBorder(registry, styleDefinition.border);
  const xfDefinition = {
    alignment: styleDefinition.alignment,
    borderId,
    fillId,
    fontId,
  };
  const key = JSON.stringify(xfDefinition);

  if (!registry.xfMap.has(key)) {
    registry.xfMap.set(key, registry.cellXfs.length);
    registry.cellXfs.push(xfDefinition);
  }

  return registry.xfMap.get(key);
}

function columnIndexToExcelName(columnIndex) {
  let currentValue = columnIndex + 1;
  let result = "";

  while (currentValue > 0) {
    const remainder = (currentValue - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    currentValue = Math.floor((currentValue - 1) / 26);
  }

  return result;
}

function buildMasterXlsxCellStyle(cellElement, rowElement, tableElement, registry, cellValue) {
  const control = cellElement.querySelector("input, select, textarea");
  const textElement = control || cellElement;
  const cellStyle = window.getComputedStyle(cellElement);
  const textStyle = window.getComputedStyle(textElement);

  return registerMasterXlsxCellStyle(registry, {
    alignment: {
      horizontal: mapHorizontalAlignment(textStyle.textAlign || cellStyle.textAlign),
      vertical: "center",
      wrapText: cellElement.tagName === "TH" || String(cellValue || "").includes("\n"),
    },
    border: {
      bottom: resolveBorderSide(cellStyle, "bottom"),
      left: resolveBorderSide(cellStyle, "left"),
      right: resolveBorderSide(cellStyle, "right"),
      top: resolveBorderSide(cellStyle, "top"),
    },
    fill: {
      color: resolveEffectiveBackgroundColor(
        [cellElement, rowElement, tableElement, dom.masterPanel, document.body],
        "FFFFFFFF"
      ),
      patternType: "solid",
    },
    font: {
      bold: Number(textStyle.fontWeight) >= 600 || cellElement.tagName === "TH",
      color: cssColorToArgb(textStyle.color) || "FF000000",
      fontName: getPrimaryFontName(textStyle.fontFamily),
      fontSize: pixelsToPoints(parseFloat(textStyle.fontSize) || 11),
      italic: textStyle.fontStyle === "italic",
    },
  });
}

function buildMasterXlsxTableModel() {
  const tableElement = dom.masterPanel?.querySelector(".master-table");

  if (!tableElement) {
    return null;
  }

  const registry = createMasterXlsxStyleRegistry();
  const rowElements = Array.from(tableElement.querySelectorAll("thead tr, tbody tr"));
  const headerCells = Array.from(tableElement.querySelectorAll("thead th"));
  const widthSourceCells = headerCells.length
    ? headerCells
    : Array.from(tableElement.querySelectorAll("tbody tr:first-child td"));
  const columns = widthSourceCells.map((cellElement) =>
    pixelsToExcelWidth(cellElement.getBoundingClientRect().width || cellElement.offsetWidth || 90)
  );
  const rows = [];
  const merges = [];
  const occupied = new Set();

  rowElements.forEach((rowElement, rowIndex) => {
    const rowNumber = rowIndex + 1;
    const cellElements = Array.from(rowElement.children).filter((child) =>
      child.tagName === "TH" || child.tagName === "TD"
    );
    let columnIndex = 0;
    const cells = [];

    cellElements.forEach((cellElement) => {
      while (occupied.has(`${rowNumber}:${columnIndex + 1}`)) {
        columnIndex += 1;
      }

      const colSpan = Math.max(1, Number(cellElement.getAttribute("colspan") || 1));
      const rowSpan = Math.max(1, Number(cellElement.getAttribute("rowspan") || 1));
      const cellValue = getMasterCellDisplayText(cellElement);
      const styleId = buildMasterXlsxCellStyle(
        cellElement,
        rowElement,
        tableElement,
        registry,
        cellValue
      );

      cells.push({
        colIndex: columnIndex,
        styleId,
        value: cellValue,
      });

      if (colSpan > 1 || rowSpan > 1) {
        const startCell = `${columnIndexToExcelName(columnIndex)}${rowNumber}`;
        const endCell = `${columnIndexToExcelName(columnIndex + colSpan - 1)}${rowNumber + rowSpan - 1}`;
        merges.push(`${startCell}:${endCell}`);
      }

      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        for (let colOffset = 0; colOffset < colSpan; colOffset += 1) {
          if (rowOffset === 0 && colOffset === 0) {
            continue;
          }

          occupied.add(`${rowNumber + rowOffset}:${columnIndex + colOffset + 1}`);
        }
      }

      columnIndex += colSpan;
    });

    rows.push({
      cells,
      height: pixelsToPoints(rowElement.getBoundingClientRect().height || rowElement.offsetHeight || 24),
      rowNumber,
    });
  });

  return {
    columns,
    merges,
    registry,
    rows,
  };
}

function buildMasterXlsxFontsXml(fonts) {
  return fonts
    .map((font) => {
      const parts = [
        "<font>",
        `<sz val="${font.fontSize}"/>`,
        `<color rgb="${font.color}"/>`,
        `<name val="${escapeXml(font.fontName)}"/>`,
        '<family val="2"/>',
      ];

      if (font.bold) {
        parts.push("<b/>");
      }

      if (font.italic) {
        parts.push("<i/>");
      }

      parts.push("</font>");
      return parts.join("");
    })
    .join("");
}

function buildMasterXlsxFillsXml(fills) {
  return fills
    .map((fill) => {
      if (fill.patternType === "none" || fill.patternType === "gray125") {
        return `<fill><patternFill patternType="${fill.patternType}"/></fill>`;
      }

      return [
        "<fill>",
        '<patternFill patternType="solid">',
        `<fgColor rgb="${fill.color}"/>`,
        '<bgColor indexed="64"/>',
        "</patternFill>",
        "</fill>",
      ].join("");
    })
    .join("");
}

function buildMasterXlsxBorderSideXml(sideName, sideDefinition) {
  if (!sideDefinition) {
    return `<${sideName}/>`;
  }

  return `<${sideName} style="${sideDefinition.style}"><color rgb="${sideDefinition.color}"/></${sideName}>`;
}

function buildMasterXlsxBordersXml(borders) {
  return borders
    .map((border) =>
      [
        "<border>",
        buildMasterXlsxBorderSideXml("left", border.left),
        buildMasterXlsxBorderSideXml("right", border.right),
        buildMasterXlsxBorderSideXml("top", border.top),
        buildMasterXlsxBorderSideXml("bottom", border.bottom),
        "<diagonal/>",
        "</border>",
      ].join("")
    )
    .join("");
}

function buildMasterXlsxAlignmentXml(alignment) {
  if (!alignment) {
    return "";
  }

  const attributes = [
    `horizontal="${alignment.horizontal}"`,
    `vertical="${alignment.vertical}"`,
  ];

  if (alignment.wrapText) {
    attributes.push('wrapText="1"');
  }

  return `<alignment ${attributes.join(" ")}/>`;
}

function buildMasterXlsxStylesXml(registry) {
  const cellXfsXml = registry.cellXfs
    .map((xf) => {
      const attributes = [
        `fontId="${xf.fontId}"`,
        `fillId="${xf.fillId}"`,
        `borderId="${xf.borderId}"`,
        'numFmtId="0"',
        'xfId="0"',
        'applyFont="1"',
        'applyFill="1"',
        'applyBorder="1"',
      ];

      if (xf.alignment) {
        attributes.push('applyAlignment="1"');
      }

      return [
        `<xf ${attributes.join(" ")}>`,
        buildMasterXlsxAlignmentXml(xf.alignment),
        "</xf>",
      ].join("");
    })
    .join("");

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<fonts count="${registry.fonts.length}">${buildMasterXlsxFontsXml(registry.fonts)}</fonts>`,
    `<fills count="${registry.fills.length}">${buildMasterXlsxFillsXml(registry.fills)}</fills>`,
    `<borders count="${registry.borders.length}">${buildMasterXlsxBordersXml(registry.borders)}</borders>`,
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    `<cellXfs count="${registry.cellXfs.length}">${cellXfsXml}</cellXfs>`,
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    "</styleSheet>",
  ].join("");
}

function buildMasterXlsxWorksheetXml(model) {
  const rowXml = model.rows
    .map((row) => {
      const cellXml = row.cells
        .map((cell) => {
          const cellReference = `${columnIndexToExcelName(cell.colIndex)}${row.rowNumber}`;
          return [
            `<c r="${cellReference}" s="${cell.styleId}" t="inlineStr">`,
            "<is>",
            `<t xml:space="preserve">${escapeXmlText(cell.value)}</t>`,
            "</is>",
            "</c>",
          ].join("");
        })
        .join("");

      return `<row r="${row.rowNumber}" ht="${row.height}" customHeight="1">${cellXml}</row>`;
    })
    .join("");
  const colXml = model.columns
    .map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`)
    .join("");
  const lastColumn = columnIndexToExcelName(Math.max(model.columns.length - 1, 0));
  const lastRow = Math.max(model.rows.length, 1);
  const mergeXml = model.merges.length
    ? `<mergeCells count="${model.merges.length}">${model.merges
        .map((mergeRef) => `<mergeCell ref="${mergeRef}"/>`)
        .join("")}</mergeCells>`
    : "";

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    `<dimension ref="A1:${lastColumn}${lastRow}"/>`,
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    `<cols>${colXml}</cols>`,
    `<sheetData>${rowXml}</sheetData>`,
    mergeXml,
    "</worksheet>",
  ].join("");
}

function buildMasterXlsxWorkbookXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="Maestro" sheetId="1" r:id="rId1"/></sheets>',
    "</workbook>",
  ].join("");
}

function buildMasterXlsxWorkbookRelsXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    "</Relationships>",
  ].join("");
}

function buildMasterXlsxRootRelsXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    "</Relationships>",
  ].join("");
}

function buildMasterXlsxContentTypesXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    "</Types>",
  ].join("");
}

function buildMasterXlsxCoreXml(timestamp) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    "<dc:creator>Cotizacion de Costura</dc:creator>",
    "<cp:lastModifiedBy>Cotizacion de Costura</cp:lastModifiedBy>",
    `<dcterms:created xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:created>`,
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:modified>`,
    "</cp:coreProperties>",
  ].join("");
}

function buildMasterXlsxAppXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    "<Application>Microsoft Excel</Application>",
    "</Properties>",
  ].join("");
}

function getMasterZipDosTimestampParts(dateValue = new Date()) {
  const year = Math.max(1980, dateValue.getFullYear());
  const month = dateValue.getMonth() + 1;
  const day = dateValue.getDate();
  const hours = dateValue.getHours();
  const minutes = dateValue.getMinutes();
  const seconds = Math.floor(dateValue.getSeconds() / 2);

  return {
    dosDate: ((year - 1980) << 9) | (month << 5) | day,
    dosTime: (hours << 11) | (minutes << 5) | seconds,
  };
}

function createCrc32Table() {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index += 1) {
    let value = index;

    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }

    table[index] = value >>> 0;
  }

  return table;
}

const MASTER_XLSX_CRC32_TABLE = createCrc32Table();

function computeCrc32(bytes) {
  let crc = 0xffffffff;

  for (let index = 0; index < bytes.length; index += 1) {
    crc = MASTER_XLSX_CRC32_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8);
  }

  return (crc ^ 0xffffffff) >>> 0;
}

function createMasterXlsxZip(files) {
  const encoder = new TextEncoder();
  const { dosDate, dosTime } = getMasterZipDosTimestampParts(new Date());
  const localParts = [];
  const centralParts = [];
  let currentOffset = 0;

  files.forEach((file) => {
    const fileNameBytes = encoder.encode(file.name);
    const fileData = typeof file.data === "string" ? encoder.encode(file.data) : file.data;
    const crc32 = computeCrc32(fileData);
    const localHeader = new Uint8Array(30 + fileNameBytes.length);
    const localView = new DataView(localHeader.buffer);

    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, dosTime, true);
    localView.setUint16(12, dosDate, true);
    localView.setUint32(14, crc32, true);
    localView.setUint32(18, fileData.length, true);
    localView.setUint32(22, fileData.length, true);
    localView.setUint16(26, fileNameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(fileNameBytes, 30);

    localParts.push(localHeader, fileData);

    const centralHeader = new Uint8Array(46 + fileNameBytes.length);
    const centralView = new DataView(centralHeader.buffer);

    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, dosTime, true);
    centralView.setUint16(14, dosDate, true);
    centralView.setUint32(16, crc32, true);
    centralView.setUint32(20, fileData.length, true);
    centralView.setUint32(24, fileData.length, true);
    centralView.setUint16(28, fileNameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, currentOffset, true);
    centralHeader.set(fileNameBytes, 46);
    centralParts.push(centralHeader);

    currentOffset += localHeader.length + fileData.length;
  });

  const centralSize = centralParts.reduce((total, part) => total + part.length, 0);
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);

  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(4, 0, true);
  endView.setUint16(6, 0, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, currentOffset, true);
  endView.setUint16(20, 0, true);

  const parts = [...localParts, ...centralParts, endRecord];
  const totalLength = parts.reduce((total, part) => total + part.length, 0);
  const zipBytes = new Uint8Array(totalLength);
  let pointer = 0;

  parts.forEach((part) => {
    zipBytes.set(part, pointer);
    pointer += part.length;
  });

  return zipBytes;
}

function buildMasterXlsxFile() {
  const tableModel = buildMasterXlsxTableModel();

  if (!tableModel) {
    return null;
  }

  const timestamp = new Date().toISOString();

  return createMasterXlsxZip([
    { name: "[Content_Types].xml", data: buildMasterXlsxContentTypesXml() },
    { name: "_rels/.rels", data: buildMasterXlsxRootRelsXml() },
    { name: "docProps/app.xml", data: buildMasterXlsxAppXml() },
    { name: "docProps/core.xml", data: buildMasterXlsxCoreXml(timestamp) },
    { name: "xl/_rels/workbook.xml.rels", data: buildMasterXlsxWorkbookRelsXml() },
    { name: "xl/styles.xml", data: buildMasterXlsxStylesXml(tableModel.registry) },
    { name: "xl/workbook.xml", data: buildMasterXlsxWorkbookXml() },
    { name: "xl/worksheets/sheet1.xml", data: buildMasterXlsxWorksheetXml(tableModel) },
  ]);
}

function downloadMasterExcelFile(fileBytes) {
  const blob = new Blob([fileBytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const fileName = getMasterExcelFileName();

  if (window.navigator?.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, fileName);
    return;
  }

  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = downloadUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();

  window.setTimeout(() => {
    URL.revokeObjectURL(downloadUrl);
  }, 1200);
}

function handleExportMasterExcel() {
  if (!appState.masterHeaders.length || appState.masterIsLoading) {
    showToast("Primero carga el Maestro de operaciones antes de exportar.", "info");
    return;
  }

  try {
    const fileBytes = buildMasterXlsxFile();

    if (!fileBytes) {
      showToast("No se pudo preparar la tabla del Maestro para exportar.", "error");
      return;
    }

    downloadMasterExcelFile(fileBytes);
    showToast("Se descargó el archivo Excel del Maestro de operaciones.", "success");
  } catch (error) {
    showToast(error.message || "No se pudo generar el archivo .xlsx del Maestro.", "error");
  }
}

function buildMasterEmptyRow() {
  const values = {};

  appState.masterHeaders.forEach((header) => {
    const headerKey = AppUtils.normalizeKey(header);
    values[headerKey] =
      getMasterComparableKey(header) === MASTER_DATE_HEADER ? getMasterCurrentDatePreview() : "";
  });

  return {
    key: buildMasterRowKey(),
    rowNumber: 0,
    values,
  };
}

function normalizeMasterSheetData(sheetData) {
  const headers = Array.isArray(sheetData && sheetData.headers) ? sheetData.headers : [];
  const rows = Array.isArray(sheetData && sheetData.rows) ? sheetData.rows : [];

  return {
    headers,
    rows: rows.map((row) => ({
      key: buildMasterRowKey(row),
      rowNumber: Number(row.rowNumber) || 0,
      values: row.values || {},
    })),
  };
}

function renderMasterTable() {
  if (!dom.masterTableHead || !dom.masterTableBody) {
    return;
  }

  renderMasterColgroup();
  dom.masterTableHead.innerHTML = "";
  dom.masterTableBody.innerHTML = "";

  if (appState.masterIsLoading) {
    const loadingRow = document.createElement("tr");
    const loadingCell = document.createElement("td");
    loadingCell.colSpan = Math.max(appState.masterHeaders.length, 1);
    loadingCell.className = "master-empty-cell";
    loadingCell.textContent = MASTER_LOADING_MESSAGE;
    loadingRow.appendChild(loadingCell);
    dom.masterTableBody.appendChild(loadingRow);
    updateMasterActionState();
    return;
  }

  if (!appState.masterHeaders.length) {
    updateMasterActionState();
    return;
  }

  const headRow = document.createElement("tr");
  appState.masterHeaders.forEach((header) => {
    const th = document.createElement("th");
    const comparableHeaderKey = getMasterComparableKey(header);
    const filterType = getMasterFilterType(comparableHeaderKey);
    const words = formatMasterHeaderLabel(header)
      .split(/\s+/)
      .filter(Boolean);

    th.dataset.header = header;
    th.dataset.comparableHeaderKey = comparableHeaderKey;

    if (filterType) {
      th.classList.add("master-filterable-header");
      th.classList.toggle("is-filtered", Boolean(getMasterFilterValue(comparableHeaderKey)));
      th.title =
        filterType === "text"
          ? "Click derecho para filtrar con texto"
          : "Click derecho para filtrar con lista";
      th.addEventListener("contextmenu", handleMasterHeaderContextMenu);
    }

    words.forEach((word) => {
      const line = document.createElement("span");
      line.className = "master-header-line";
      line.textContent = word;
      th.appendChild(line);
    });

    headRow.appendChild(th);
  });
  dom.masterTableHead.appendChild(headRow);

  if (!appState.masterRows.length) {
    const emptyRow = document.createElement("tr");
    const emptyCell = document.createElement("td");
    emptyCell.colSpan = appState.masterHeaders.length;
    emptyCell.className = "master-empty-cell";
    emptyCell.textContent = appState.masterSourceRows.length
      ? "No hay filas que coincidan con los filtros activos."
      : "No hay registros en basedatos. Usa Agregar para crear la primera fila.";
    emptyRow.appendChild(emptyCell);
    dom.masterTableBody.appendChild(emptyRow);
    updateMasterActionState();
    return;
  }

  appState.masterRows.forEach((rowData) => {
    const rowElement = document.createElement("tr");
    const isEditing = rowData.key === appState.masterEditingRowKey;

    rowElement.className = "master-row";
    rowElement.dataset.rowKey = rowData.key;
    if (isEditing) {
      rowElement.classList.add("is-editing");
    }

    appState.masterHeaders.forEach((header) => {
      const headerKey = AppUtils.normalizeKey(header);
      const comparableHeaderKey = getMasterComparableKey(header);
      const cell = document.createElement("td");
      const isDateCell = comparableHeaderKey === MASTER_DATE_HEADER;
      const isEditableCell = isEditing && !isDateCell;
      
      const dropdownColumns = new Set(["SECCION", "BLOQUE", "TIPO_OPERACION", "ACCION", "TIPO_MAQ", "MAQUINA", "ESTADO"]);
      const useDropdown = isEditableCell && dropdownColumns.has(comparableHeaderKey);
      
      const input = document.createElement(useDropdown ? "select" : "input");
      
      if (!useDropdown) {
        input.type = "text";
        input.readOnly = !isEditableCell;
      } else {
        const defaultOption = document.createElement("option");
        defaultOption.value = "";
        defaultOption.textContent = "Seleccionar";
        input.appendChild(defaultOption);
        
        getMasterUniqueFilterValues(comparableHeaderKey).forEach(val => {
          if (val) {
            const option = document.createElement("option");
            option.value = val;
            option.textContent = val;
            input.appendChild(option);
          }
        });
      }

      input.className = "table-input master-table-input";
      input.value = rowData.values[headerKey] || "";
      input.dataset.rowKey = rowData.key;
      input.dataset.headerKey = headerKey;
      input.classList.toggle("center-cell", isMasterCenteredColumn(comparableHeaderKey));

      input.tabIndex = isEditableCell ? 0 : -1;

      if (isDateCell) {
        const rawDateValue = rowData.values[headerKey] || "";
        const dateTooltip = formatMasterDateTooltip(rawDateValue);

        input.value = formatMasterDateShortValue(rawDateValue);
        input.title = dateTooltip;
        input.classList.toggle("has-tooltip", Boolean(dateTooltip));
      }

      if (isMasterFullContentTooltipColumn(comparableHeaderKey)) {
        const tooltipValue = String(rowData.values[headerKey] || "").trim();
        input.title = tooltipValue;
        input.classList.toggle("has-tooltip", Boolean(tooltipValue));
      }

      if (comparableHeaderKey === MASTER_EDIT_TRIGGER_HEADER) {
        input.addEventListener("contextmenu", handleMasterEditTrigger);
      }

      cell.appendChild(input);
      rowElement.appendChild(cell);
    });

    dom.masterTableBody.appendChild(rowElement);
  });

  updateMasterActionState();
}

function focusMasterEditingRow() {
  if (!appState.masterEditingRowKey || !dom.masterTableBody) {
    return;
  }

  const firstEditableCell = dom.masterTableBody.querySelector(
    `.master-table-input[data-row-key="${appState.masterEditingRowKey}"]:not([readonly])`
  );

  if (!firstEditableCell) {
    return;
  }

  firstEditableCell.focus();
  firstEditableCell.select();
  firstEditableCell.closest("tr")?.scrollIntoView({
    block: "nearest",
    inline: "nearest",
  });
}

function setMasterRows(sheetData) {
  const normalizedData = normalizeMasterSheetData(sheetData);
  const availableComparableHeaders = new Set(
    normalizedData.headers.map((header) => getMasterComparableKey(header))
  );

  appState.masterHeaders = normalizedData.headers;
  appState.masterSourceRows = normalizedData.rows;
  appState.masterFilters = Object.entries(appState.masterFilters).reduce((accumulator, [key, value]) => {
    if (availableComparableHeaders.has(key) && getMasterFilterType(key) && String(value || "").trim()) {
      accumulator[key] = String(value).trim();
    }

    return accumulator;
  }, {});
  appState.masterEditingRowKey = null;
  closeMasterFilterMenu();
  applyMasterFilters();
  renderMasterTable();
  updateMasterLoadedSummary();
}

async function loadMasterData(options = {}) {
  const { keepView = false } = options;

  if (!AppsScriptAPI.isConfigured()) {
    throw new Error("Primero configura WEB_APP_URL para usar el Maestro de operaciones.");
  }

  setMasterLoadingState(true);
  updateMasterInfo(MASTER_LOADING_MESSAGE);

  try {
    const response = await AppsScriptAPI.fetchBasedatosSheet();
    const payload = response.data || response;

    if (!response.success && !Array.isArray(payload.rows)) {
      throw new Error(response.message || "No se pudo cargar la hoja basedatos.");
    }

    appState.masterIsLoading = false;
    setMasterRows(payload);
  } catch (error) {
    setMasterLoadingState(false);
    throw error;
  }

  if (!keepView) {
    switchView("master");
  }

  updateMasterLoadedSummary();
}

function beginMasterRowEdit(rowKey) {
  if (appState.masterEditingRowKey && appState.masterEditingRowKey !== rowKey) {
    showToast("Guarda la fila actual antes de editar otra en el Maestro de operaciones.", "info");
    focusMasterEditingRow();
    return;
  }

  if (appState.masterEditingRowKey === rowKey) {
    focusMasterEditingRow();
    return;
  }

  appState.masterEditingRowKey = rowKey;
  applyMasterFilters();
  renderMasterTable();
  updateMasterInfo("Edita la fila seleccionada y luego guarda para reemplazarla al final del sheet.");
  focusMasterEditingRow();
}

function handleMasterEditTrigger(event) {
  if (appState.currentView !== "master") {
    return;
  }

  event.preventDefault();
  beginMasterRowEdit(event.currentTarget.dataset.rowKey);
}

async function handleOpenMasterView() {
  switchView("master");

  if (!AppsScriptAPI.isConfigured()) {
    updateMasterInfo("Configura WEB_APP_URL para cargar la hoja basedatos.");
    showToast("Actualiza WEB_APP_URL antes de abrir el Maestro de operaciones.", "error");
    return;
  }

  try {
    await loadMasterData({ keepView: true });
  } catch (error) {
    updateMasterInfo("No se pudo cargar la hoja basedatos.");
    showToast(error.message || "No se pudo cargar el Maestro de operaciones.", "error");
  }
}

function handleAddMasterRow() {
  if (!appState.masterHeaders.length) {
    showToast("Primero carga el Maestro de operaciones.", "info");
    return;
  }

  if (appState.masterEditingRowKey) {
    showToast("Guarda la fila actual antes de agregar otra.", "info");
    focusMasterEditingRow();
    return;
  }

  appState.masterSourceRows = [...appState.masterSourceRows, buildMasterEmptyRow()];
  appState.masterEditingRowKey = appState.masterSourceRows[appState.masterSourceRows.length - 1].key;
  applyMasterFilters();
  renderMasterTable();
  updateMasterInfo("Completa todas las columnas y luego guarda la nueva fila.");
  focusMasterEditingRow();
}

function collectMasterEditingRowData() {
  const editingRow = getMasterRowByKey(appState.masterEditingRowKey);

  if (!editingRow || !dom.masterTableBody) {
    return null;
  }

  const rowInputs = dom.masterTableBody.querySelectorAll(
    `.master-table-input[data-row-key="${appState.masterEditingRowKey}"]`
  );
  const rowData = {};

  rowInputs.forEach((input) => {
    rowData[input.dataset.headerKey] = input.value.trim();
  });

  return {
    editingRow,
    rowData,
  };
}

function validateMasterRowData(rowData) {
  for (const header of appState.masterHeaders) {
    const headerKey = AppUtils.normalizeKey(header);

    if (getMasterComparableKey(header) === MASTER_DATE_HEADER) {
      continue;
    }

    if (!String(rowData[headerKey] || "").trim()) {
      return {
        isValid: false,
        focusElement: dom.masterTableBody.querySelector(
          `.master-table-input[data-row-key="${appState.masterEditingRowKey}"][data-header-key="${headerKey}"]`
        ),
      };
    }
  }

  return {
    isValid: true,
    focusElement: null,
  };
}

async function handleSaveMasterRow() {
  if (!appState.masterEditingRowKey) {
    showToast("No hay ninguna fila en edición en el Maestro de operaciones.", "info");
    return;
  }

  if (!AppsScriptAPI.isConfigured()) {
    showToast("Actualiza WEB_APP_URL antes de guardar en basedatos.", "error");
    return;
  }

  const collectedData = collectMasterEditingRowData();

  if (!collectedData) {
    showToast("No se pudo leer la fila en edición.", "error");
    return;
  }

  const validation = validateMasterRowData(collectedData.rowData);

  if (!validation.isValid) {
    showToast("Debes completar todas las columnas antes de guardar.", "error");
    focusFormControl(validation.focusElement);
    return;
  }

  dom.masterSaveBtn.disabled = true;
  dom.masterAddBtn.disabled = true;
  dom.masterSaveBtn.classList.add("is-saving");
  updateMasterInfo("Guardando fila en basedatos...");

  try {
    const response = await AppsScriptAPI.saveBasedatosRow({
      originalRowNumber: collectedData.editingRow.rowNumber || "",
      rowData: collectedData.rowData,
    });

    if (!response.success) {
      throw new Error(response.message || "No se pudo guardar la fila en basedatos.");
    }

    await refreshCatalogData({ notifySuccess: false });
    await loadMasterData({ keepView: true });
    showToast(response.message || "Fila guardada correctamente en basedatos.", "success");
  } catch (error) {
    updateMasterInfo("No se pudo guardar la fila en basedatos.");
    showToast(error.message || "Ocurrió un error al guardar en basedatos.", "error");
  } finally {
    dom.masterSaveBtn.classList.remove("is-saving");
    updateMasterActionState();
  }
}

function handleSearchProtoInput(event) {
  const normalizedValue = normalizeSearchProto(event.target.value);

  if (event.target.value !== normalizedValue) {
    event.target.value = normalizedValue;
  }
}

function initializeClientField() {
  renderClientOptions();
  hideCustomClientInput();
  updateDeleteClientButton();
}

function getStoredCustomClients() {
  try {
    const rawValue = window.localStorage.getItem(CUSTOM_CLIENTS_STORAGE_KEY);
    const parsed = JSON.parse(rawValue || "[]");

    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .map((item) => String(item || "").trim())
      .filter(Boolean);
  } catch (error) {
    return [];
  }
}

function saveStoredCustomClients(clientList) {
  try {
    window.localStorage.setItem(CUSTOM_CLIENTS_STORAGE_KEY, JSON.stringify(clientList));
  } catch (error) {
    // Ignore storage errors and keep the new client in the current session.
  }
}

function getAllClientOptions() {
  const options = [...DEFAULT_CLIENT_OPTIONS, ...getStoredCustomClients()];
  const unique = [];
  const seen = new Set();

  options.forEach((item) => {
    const cleanItem = String(item || "").trim();
    const normalized = AppUtils.normalizeKey(cleanItem);

    if (!cleanItem || seen.has(normalized)) {
      return;
    }

    seen.add(normalized);
    unique.push(cleanItem);
  });

  return unique;
}

function renderClientOptions(selectedValue = "") {
  const options = getAllClientOptions();
  const normalizedSelected = AppUtils.normalizeKey(selectedValue);

  dom.form.cliente.innerHTML = "";

  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = "Seleccionar";
  dom.form.cliente.appendChild(placeholderOption);

  options.forEach((clientName) => {
    const option = document.createElement("option");
    option.value = clientName;
    option.textContent = clientName;
    dom.form.cliente.appendChild(option);
  });

  const newOption = document.createElement("option");
  newOption.value = NEW_CLIENT_OPTION_VALUE;
  newOption.textContent = "NUEVO";
  dom.form.cliente.appendChild(newOption);

  const match = options.find((item) => AppUtils.normalizeKey(item) === normalizedSelected);
  dom.form.cliente.value = match || "";
  updateDeleteClientButton();
}

function ensureClientOption(clientName, persist = false) {
  const cleanName = normalizeUppercaseTextValue(clientName).trim();

  if (!cleanName) {
    return "";
  }

  const options = getAllClientOptions();
  const normalizedName = AppUtils.normalizeKey(cleanName);
  const existingName = options.find((item) => AppUtils.normalizeKey(item) === normalizedName);

  if (existingName) {
    return existingName;
  }

  if (persist) {
    const updatedClients = [...getStoredCustomClients(), cleanName];
    saveStoredCustomClients(updatedClients);
  }

  return cleanName;
}

function setClientValue(clientName) {
  const cleanName = normalizeUppercaseTextValue(clientName).trim();

  if (!cleanName) {
    renderClientOptions();
    hideCustomClientInput();
    updateDeleteClientButton();
    return;
  }

  const resolvedName = ensureClientOption(cleanName, true);
  renderClientOptions(resolvedName);
  dom.form.cliente.value = resolvedName;
  hideCustomClientInput();
  updateDeleteClientButton();
}

function handleClientSelectChange() {
  if (dom.form.cliente.disabled) {
    return;
  }

  if (dom.form.cliente.value === NEW_CLIENT_OPTION_VALUE) {
    updateDeleteClientButton();
    showCustomClientInput();
    return;
  }

  hideCustomClientInput();
  updateDeleteClientButton();
}

function showCustomClientInput() {
  dom.form.clienteCustom.classList.remove("hidden");
  dom.form.clienteCustom.value = "";
  dom.form.clienteCustom.focus();
}

function hideCustomClientInput() {
  dom.form.clienteCustom.classList.add("hidden");
  dom.form.clienteCustom.value = "";
}

function isDefaultClient(clientName) {
  const normalizedName = AppUtils.normalizeKey(clientName);
  return DEFAULT_CLIENT_OPTIONS.some((item) => AppUtils.normalizeKey(item) === normalizedName);
}

function isStoredCustomClient(clientName) {
  const normalizedName = AppUtils.normalizeKey(clientName);
  return getStoredCustomClients().some((item) => AppUtils.normalizeKey(item) === normalizedName);
}

function updateDeleteClientButton() {
  const selectedClient = dom.form.cliente.value;
  const canDelete =
    isEditableMode() &&
    Boolean(selectedClient) &&
    selectedClient !== NEW_CLIENT_OPTION_VALUE &&
    !isDefaultClient(selectedClient) &&
    isStoredCustomClient(selectedClient);

  dom.form.clienteDelete.classList.toggle("hidden", !canDelete);
}

function commitCustomClient() {
  if (!isEditableMode()) {
    return;
  }
  if (dom.form.cliente.value !== NEW_CLIENT_OPTION_VALUE) {
    return;
  }

  const newClientName = dom.form.clienteCustom.value.trim();

  if (!newClientName) {
    return;
  }

  const resolvedName = ensureClientOption(newClientName, true);
  renderClientOptions(resolvedName);
  dom.form.cliente.value = resolvedName;
  hideCustomClientInput();
  updateDeleteClientButton();
  showToast(`Cliente ${resolvedName} agregado a la lista.`, "success");
}

function deleteSelectedCustomClient() {
  if (!isEditableMode()) {
    return;
  }
  const selectedClient = dom.form.cliente.value;

  if (
    !selectedClient ||
    selectedClient === NEW_CLIENT_OPTION_VALUE ||
    isDefaultClient(selectedClient) ||
    !isStoredCustomClient(selectedClient)
  ) {
    return;
  }

  if (!window.confirm(`Se eliminarÃ¡ el cliente ${selectedClient} de la lista local. Â¿Deseas continuar?`)) {
    return;
  }

  const normalizedTarget = AppUtils.normalizeKey(selectedClient);
  const updatedClients = getStoredCustomClients().filter(
    (item) => AppUtils.normalizeKey(item) !== normalizedTarget
  );

  saveStoredCustomClients(updatedClients);
  renderClientOptions();
  hideCustomClientInput();
  updateDeleteClientButton();
  showToast(`Cliente ${selectedClient} eliminado de la lista.`, "success");
}

// --- Tab Logic ---
function handleTabSwitch(event) {
  const tabBtn = event.currentTarget;
  if (tabBtn.classList.contains("is-active")) return;
  
  [dom.tabCostura, dom.tabCorte, dom.tabAcabado, dom.tabResumen].forEach(t => t?.classList.remove("is-active"));
  tabBtn.classList.add("is-active");
  
  dom.costuraPanel.classList.add("hidden");
  dom.cortePanel.classList.add("hidden");
  dom.acabadoPanel.classList.add("hidden");
  dom.resumenPanel.classList.add("hidden");
  
  const targetId = tabBtn.dataset.target;
  if (dom[targetId]) dom[targetId].classList.remove("hidden");
}

function buildInitialCorteRows() {
  if (!dom.corteTableBody) return;
  dom.corteTableBody.innerHTML = "";
  APP_CONFIG.defaultCorteOperations.forEach(item => {
    appendCorteRow({ 
      operaciones: item.operaciones,
      area: item.area
    });
  });
  refreshCorteSummary();
}

function buildInitialAcabadoRows() {
  if (!dom.acabadoTableBody) return;
  dom.acabadoTableBody.innerHTML = "";
  APP_CONFIG.defaultAcabadoOperations.forEach((operaciones) => {
    appendAcabadoRow({ operaciones });
  });
  refreshAcabadoSummary();
}

// --- Corte Logic ---
function createCorteRow(rowData = null, options = {}) {
  const fragment = dom.corteRowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  const inputs = Array.from(row.querySelectorAll('.table-input'));
  row.dataset.hideZeroValues = options.hideZeroValues ? "true" : "false";
  
  inputs.forEach(input => {
    if (['tiempoEstimadoCorte', 'tiempoEstimadoHabilitado'].includes(input.dataset.field)) {
      input.addEventListener("focus", (e) => {
        if (!isEditableMode()) return;
        e.target.value = e.target.dataset.formula !== undefined ? e.target.dataset.formula : e.target.value;
      });
      input.addEventListener("blur", (e) => {
        if (!isEditableMode()) return;
        const rawValue = e.target.value.trim();
        e.target.dataset.formula = rawValue;
        e.target.value = rawValue ? AppUtils.formatNumber(AppUtils.evaluateFormula(rawValue), 2) : "";
        updateCorteTotals(row);
      });
      input.addEventListener("input", (e) => {
        e.target.dataset.formula = e.target.value.trim();
        updateCorteTotals(row);
      });
    } else if (input.dataset.field === 'proteccion') {
      input.addEventListener("dblclick", handleProtectionDoubleClick);
      input.addEventListener("input", handleProtectionInput);
      input.addEventListener("blur", handleProtectionBlur);
      input.addEventListener("keydown", handleProtectionKeydown);
    } else if (input.dataset.field === 'area') {
      input.addEventListener("input", () => updateCorteTotals(row));
    } else if (input.dataset.field === 'operaciones') {
      input.addEventListener("contextmenu", (event) => handleOperationsContextMenu(event, "corte"));
    }
  });

  writeCorteRowValues(row, buildEmptyCorteRow());
  if (rowData) applyCorteRowData(row, rowData);
  syncRowInteractivity(row);
  return row;
}

function appendCorteRow(rowData = null, options = {}) {
  const row = createCorteRow(rowData, options);
  dom.corteTableBody.appendChild(row);
  return row;
}

function removeLastCorteRow() {
  if (!isEditableMode()) return;
  const rows = Array.from(dom.corteTableBody.querySelectorAll("tr"));
  if (rows.length <= 1) {
    if (rows[0]) writeCorteRowValues(rows[0], buildEmptyCorteRow());
  } else {
    rows[rows.length - 1].remove();
  }
  refreshCorteSummary();
}

function clearCorteTableWithConfirm() {
  if (!isEditableMode()) return;
  openConfirmModal("Esta seguro que quiere eliminar todos los datos de Corte?", () => {
    buildInitialCorteRows();
    applyTableInteractivity();
  }, { acceptLabel: "Eliminar" });
}

function buildEmptyCorteRow() {
  return {
    operaciones: "",
    tiempoEstimadoCorte: "",
    tiempoEstimadoHabilitado: "",
    proteccion: appState.corteProtectionValue,
    area: "",
    tiempoCorte: "",
    tiempoHab: "",
    tiempoCotizacion: "",
  };
}

function writeCorteRowValues(row, rowData) {
  const fields = row.querySelectorAll("[data-field]");
  const hideZeroValues = row.dataset.hideZeroValues === "true";
  fields.forEach((field) => {
    const fieldName = field.dataset.field;
    if (["tiempoEstimadoCorte", "tiempoEstimadoHabilitado"].includes(fieldName)) {
      const rawValue = rowData[fieldName] !== undefined ? String(rowData[fieldName]) : "";
      field.dataset.formula = rawValue;
      
      // Do not overwrite the value while the user is actively typing in it
      if (document.activeElement !== field) {
        const evaluatedValue = rawValue ? AppUtils.evaluateFormula(rawValue) : "";
        field.value = rawValue && !(hideZeroValues && AppUtils.safeNumber(evaluatedValue) === 0)
          ? AppUtils.formatNumber(evaluatedValue, 2)
          : "";
      }
    } else if (["tiempoCorte", "tiempoHab", "tiempoCotizacion"].includes(fieldName)) {
      const numericValue = AppUtils.safeNumber(rowData[fieldName]);
      if (rowData[fieldName] === "" || (hideZeroValues && numericValue === 0)) {
        field.value = "";
      } else {
        field.value = AppUtils.formatNumber(numericValue, 2);
      }
    } else if (fieldName === "proteccion") {
      const numericValue = AppUtils.safeNumber(rowData[fieldName]);
      field.value = hideZeroValues && numericValue === 0 ? "" : AppUtils.formatNumber(numericValue, 2);
    } else {
      if (document.activeElement !== field) {
        field.value = rowData[fieldName] || "";
      }
    }
  });
}

function readCorteRowValues(row) {
  const getValue = (fieldName) => row.querySelector(`[data-field="${fieldName}"]`).value;
  const getFormula = (fieldName) => {
    const el = row.querySelector(`[data-field="${fieldName}"]`);
    return el.dataset.formula !== undefined ? el.dataset.formula : el.value;
  };
  const tiempoCorteValue = getValue("tiempoCorte").trim();
  const tiempoHabValue = getValue("tiempoHab").trim();
  const tiempoCotizacionValue = getValue("tiempoCotizacion").trim();

  return {
    operaciones: getValue("operaciones").trim(),
    tiempoEstimadoCorte: getFormula("tiempoEstimadoCorte").trim(),
    tiempoEstimadoHabilitado: getFormula("tiempoEstimadoHabilitado").trim(),
    proteccion: AppUtils.safeNumber(getValue("proteccion")) || APP_CONFIG.defaultProtection,
    area: getValue("area").trim(),
    tiempoCorte: tiempoCorteValue ? AppUtils.safeNumber(tiempoCorteValue) : "",
    tiempoHab: tiempoHabValue ? AppUtils.safeNumber(tiempoHabValue) : "",
    tiempoCotizacion: tiempoCotizacionValue ? AppUtils.safeNumber(tiempoCotizacionValue) : "",
  };
}

function applyCorteRowData(row, rowData) {
  writeCorteRowValues(row, {
    operaciones: rowData.operaciones || "",
    tiempoEstimadoCorte: rowData.tiempoEstimadoCorte !== undefined ? String(rowData.tiempoEstimadoCorte) : "",
    tiempoEstimadoHabilitado: rowData.tiempoEstimadoHabilitado !== undefined ? String(rowData.tiempoEstimadoHabilitado) : "",
    proteccion: AppUtils.safeNumber(rowData.proteccion || APP_CONFIG.defaultProtection),
    area: rowData.area || "",
    tiempoCorte:
      rowData.tiempoCorte === "" || rowData.tiempoCorte === null || rowData.tiempoCorte === undefined
        ? ""
        : AppUtils.safeNumber(rowData.tiempoCorte),
    tiempoHab:
      rowData.tiempoHab === "" || rowData.tiempoHab === null || rowData.tiempoHab === undefined
        ? ""
        : AppUtils.safeNumber(rowData.tiempoHab),
    tiempoCotizacion:
      rowData.tiempoCotizacion === "" || rowData.tiempoCotizacion === null || rowData.tiempoCotizacion === undefined
        ? ""
        : AppUtils.safeNumber(rowData.tiempoCotizacion),
  });
}

function updateCorteTotals(row) {
  const data = readCorteRowValues(row);
  const tcExt = AppUtils.evaluateFormula(data.tiempoEstimadoCorte);
  const thExt = AppUtils.evaluateFormula(data.tiempoEstimadoHabilitado);
  
  const isCorte = data.area.toUpperCase() === "CORT";
  const isHab = data.area.toUpperCase() === "HAB";
  const hasCorteEstimate = Boolean(data.tiempoEstimadoCorte);
  const hasHabEstimate = Boolean(data.tiempoEstimadoHabilitado);

  const tiempoCorte = isCorte && hasCorteEstimate ? (tcExt * data.proteccion) : "";
  const tiempoHab = isHab && hasHabEstimate ? (thExt * data.proteccion) : "";
  
  const valCorte = AppUtils.safeNumber(tiempoCorte);
  const valHab = AppUtils.safeNumber(tiempoHab);
  const tiempoCotizacion = tiempoCorte === "" && tiempoHab === "" ? "" : (valCorte + valHab);
  
  writeCorteRowValues(row, { ...data, tiempoCorte, tiempoHab, tiempoCotizacion });
  refreshCorteSummary();
}

function getCorteFilledRows() {
  return Array.from(dom.corteTableBody.querySelectorAll("tr")).map(readCorteRowValues).filter(r => r.operaciones || r.tiempoEstimadoCorte || r.tiempoEstimadoHabilitado);
}

function refreshCorteSummary() {
  const rows = Array.from(dom.corteTableBody?.querySelectorAll("tr") || []);
  let tcorte = 0, thab = 0, trestCorte = 0, trestHab = 0, tcot = 0;
  rows.forEach(row => {
    const data = readCorteRowValues(row);
    trestCorte += AppUtils.evaluateFormula(data.tiempoEstimadoCorte); 
    trestHab += AppUtils.evaluateFormula(data.tiempoEstimadoHabilitado);
    tcorte += AppUtils.safeNumber(data.tiempoCorte);
    thab += AppUtils.safeNumber(data.tiempoHab);
    tcot += AppUtils.safeNumber(data.tiempoCotizacion);
  });
  if (dom.summaryCorte) {
    dom.summaryCorte.estimatedCorte.textContent = AppUtils.formatNumber(trestCorte, 2);
    dom.summaryCorte.estimatedHab.textContent = AppUtils.formatNumber(trestHab, 2);
    dom.summaryCorte.corte.textContent = AppUtils.formatNumber(tcorte, 2);
    dom.summaryCorte.hab.textContent = AppUtils.formatNumber(thab, 2);
    dom.summaryCorte.cotizacion.textContent = AppUtils.formatNumber(tcot, 2);
  }
  refreshResumenSummary();
}

function renderCorteRows(rows, options = {}) {
  appState.corteProtectionValue = resolveActiveProtectionValue(rows, APP_CONFIG.defaultProtection);
  if (dom.corteTableBody) dom.corteTableBody.innerHTML = "";
  if (rows.length) {
    rows.forEach(r => appendCorteRow(r, options));
  } else {
    appendCorteRow(null, options);
  }
  refreshCorteSummary();
}

// --- Acabado Logic ---
function createAcabadoRow(rowData = null) {
  const fragment = dom.acabadoRowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  const inputs = Array.from(row.querySelectorAll('.table-input'));
  
  inputs.forEach(input => {
    if (input.dataset.field === "operaciones") {
      input.addEventListener("contextmenu", (event) => handleOperationsContextMenu(event, "acabado"));
    } else if (input.dataset.field === "tiempoEstimado") {
      input.addEventListener("focus", (event) => {
        if (!isEditableMode()) return;
        const rawValue = event.target.value.trim();
        event.target.value = rawValue ? String(roundToDecimals(AppUtils.safeNumber(rawValue), 2)) : "";
        event.target.select();
      });
      input.addEventListener("input", (event) => {
        if (!isEditableMode()) return;
        const normalizedValue = normalizeDecimalInputValue(event.target.value, 2);
        if (event.target.value !== normalizedValue) {
          event.target.value = normalizedValue;
        }
        updateAcabadoTotals(row);
      });
      input.addEventListener("blur", (event) => {
        if (!isEditableMode()) return;
        const normalizedValue = normalizeDecimalInputValue(event.target.value, 2);
        const hasDigits = /\d/.test(normalizedValue);
        event.target.value = hasDigits
          ? AppUtils.formatNumber(AppUtils.safeNumber(normalizedValue), 2)
          : "";
        updateAcabadoTotals(row);
      });
      input.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") return;
        event.preventDefault();
        event.target.blur();
      });
    } else if (input.dataset.field === "proteccion") {
      input.addEventListener("dblclick", handleProtectionDoubleClick);
      input.addEventListener("input", handleProtectionInput);
      input.addEventListener("blur", handleProtectionBlur);
      input.addEventListener("keydown", handleProtectionKeydown);
    }
  });

  writeAcabadoRowValues(row, buildEmptyAcabadoRow());
  if (rowData) applyAcabadoRowData(row, rowData);
  syncRowInteractivity(row);
  return row;
}

function appendAcabadoRow(rowData = null) {
  const row = createAcabadoRow(rowData);
  dom.acabadoTableBody.appendChild(row);
  return row;
}

function removeLastAcabadoRow() {
  if (!isEditableMode()) return;
  const rows = Array.from(dom.acabadoTableBody.querySelectorAll("tr"));
  if (rows.length <= 1) {
    if (rows[0]) writeAcabadoRowValues(rows[0], buildEmptyAcabadoRow());
  } else {
    rows[rows.length - 1].remove();
  }
  refreshAcabadoSummary();
}

function clearAcabadoTableWithConfirm() {
  if (!isEditableMode()) return;
  openConfirmModal("Esta seguro que quiere eliminar todos los datos de Acabados?", () => {
    buildInitialAcabadoRows();
    refreshAcabadoSummary();
    applyTableInteractivity();
  }, { acceptLabel: "Eliminar" });
}

function getCorteRows() {
  return Array.from(dom.corteTableBody?.querySelectorAll("tr") || []);
}

function getAcabadoRows() {
  return Array.from(dom.acabadoTableBody?.querySelectorAll("tr") || []);
}

function insertCorteRowAt(index, rowData = null) {
  const rows = getCorteRows();
  const row = createCorteRow(rowData);
  syncRowInteractivity(row);
  if (index >= rows.length) {
    dom.corteTableBody.appendChild(row);
  } else {
    dom.corteTableBody.insertBefore(row, rows[index]);
  }
  refreshCorteSummary();
  return row;
}

function insertAcabadoRowAt(index, rowData = null) {
  const rows = getAcabadoRows();
  const row = createAcabadoRow(rowData);
  syncRowInteractivity(row);
  if (index >= rows.length) {
    dom.acabadoTableBody.appendChild(row);
  } else {
    dom.acabadoTableBody.insertBefore(row, rows[index]);
  }
  refreshAcabadoSummary();
  return row;
}

function insertCorteRowAbove(referenceRow) {
  if (!isEditableMode()) return;
  const index = getCorteRows().indexOf(referenceRow);
  if (index === -1) return;
  const inserted = insertCorteRowAt(index);
  inserted.querySelector('[data-field="operaciones"]')?.focus();
}

function insertCorteRowBelow(referenceRow) {
  if (!isEditableMode()) return;
  const index = getCorteRows().indexOf(referenceRow);
  if (index === -1) return;
  const inserted = insertCorteRowAt(index + 1);
  inserted.querySelector('[data-field="operaciones"]')?.focus();
}

function insertAcabadoRowAbove(referenceRow) {
  if (!isEditableMode()) return;
  const index = getAcabadoRows().indexOf(referenceRow);
  if (index === -1) return;
  const inserted = insertAcabadoRowAt(index);
  inserted.querySelector('[data-field="operaciones"]')?.focus();
}

function insertAcabadoRowBelow(referenceRow) {
  if (!isEditableMode()) return;
  const index = getAcabadoRows().indexOf(referenceRow);
  if (index === -1) return;
  const inserted = insertAcabadoRowAt(index + 1);
  inserted.querySelector('[data-field="operaciones"]')?.focus();
}

function deleteCorteTableRow(rowToRemove) {
  if (!isEditableMode() || !rowToRemove) return;
  const rows = getCorteRows();
  if (rows.length === 1) {
    writeCorteRowValues(rows[0], buildEmptyCorteRow());
    refreshCorteSummary();
    return;
  }
  const nextFocusRow = rowToRemove.nextElementSibling || rowToRemove.previousElementSibling;
  rowToRemove.remove();
  refreshCorteSummary();
  nextFocusRow?.querySelector('[data-field="operaciones"]')?.focus();
}

function deleteAcabadoTableRow(rowToRemove) {
  if (!isEditableMode() || !rowToRemove) return;
  const rows = getAcabadoRows();
  if (rows.length === 1) {
    writeAcabadoRowValues(rows[0], buildEmptyAcabadoRow());
    rows[0].querySelector('[data-field="operaciones"]')?.focus();
    refreshAcabadoSummary();
    return;
  }
  const nextFocusRow = rowToRemove.nextElementSibling || rowToRemove.previousElementSibling;
  rowToRemove.remove();
  refreshAcabadoSummary();
  nextFocusRow?.querySelector('[data-field="operaciones"]')?.focus();
}

function removeLastCorteRow() {
  if (!isEditableMode()) return;
  const rows = getCorteRows();
  if (rows.length <= 1) {
    if (rows[0]) writeCorteRowValues(rows[0], buildEmptyCorteRow());
  } else {
    rows[rows.length - 1].remove();
  }
  refreshCorteSummary();
}

function clearCorteTableWithConfirm() {
  if (!isEditableMode()) return;
  openConfirmModal("Esta seguro que quiere eliminar todos los datos de Corte?", () => {
    buildInitialCorteRows();
    applyTableInteractivity();
  }, { acceptLabel: "Eliminar" });
}


function buildEmptyAcabadoRow() {
  return {
    operaciones: "",
    tiempoEstimado: "",
    proteccion: appState.acabadoProtectionValue,
    tiempoCotizacion: "",
  };
}

function writeAcabadoRowValues(row, rowData) {
  const fields = row.querySelectorAll("[data-field]");
  fields.forEach((field) => {
    const fieldName = field.dataset.field;
    if (fieldName === "tiempoEstimado") {
      if (document.activeElement !== field) {
        field.value = rowData[fieldName] === "" ? "" : AppUtils.formatNumber(rowData[fieldName] || 0, 2);
      }
    } else if (fieldName === "proteccion") {
      field.value = AppUtils.formatNumber(rowData[fieldName] || 0, 2);
    } else if (fieldName === "tiempoCotizacion") {
      field.value = rowData[fieldName] === "" ? "" : AppUtils.formatNumber(rowData[fieldName] || 0, 2);
    } else if (document.activeElement !== field) {
      field.value = rowData[fieldName] || "";
    }
  });
}

function readAcabadoRowValues(row) {
  const getValue = (fieldName) => row.querySelector(`[data-field="${fieldName}"]`).value;
  const tiempoEstimadoValue = getValue("tiempoEstimado").trim();
  const tiempoCotizacionValue = getValue("tiempoCotizacion").trim();
  return {
    operaciones: getValue("operaciones").trim(),
    tiempoEstimado: tiempoEstimadoValue ? AppUtils.safeNumber(tiempoEstimadoValue) : "",
    proteccion: AppUtils.safeNumber(getValue("proteccion")) || APP_CONFIG.defaultProtection,
    tiempoCotizacion: tiempoCotizacionValue ? AppUtils.safeNumber(tiempoCotizacionValue) : "",
  };
}

function applyAcabadoRowData(row, rowData) {
  writeAcabadoRowValues(row, {
    operaciones: rowData.operaciones || "",
    tiempoEstimado:
      rowData.tiempoEstimado === "" || rowData.tiempoEstimado === null || rowData.tiempoEstimado === undefined
        ? ""
        : AppUtils.safeNumber(rowData.tiempoEstimado),
    proteccion: AppUtils.safeNumber(rowData.proteccion || APP_CONFIG.defaultProtection),
    tiempoCotizacion:
      rowData.tiempoCotizacion === "" || rowData.tiempoCotizacion === null || rowData.tiempoCotizacion === undefined
        ? ""
        : AppUtils.safeNumber(rowData.tiempoCotizacion),
  });
}

function updateAcabadoTotals(row) {
  const data = readAcabadoRowValues(row);
  const tiempoCotizacion = data.tiempoEstimado === "" ? "" : data.tiempoEstimado * data.proteccion;
  writeAcabadoRowValues(row, { ...data, tiempoCotizacion });
  refreshAcabadoSummary();
}

function getAcabadoFilledRows() {
  return Array.from(dom.acabadoTableBody.querySelectorAll("tr")).map(readAcabadoRowValues).filter(r => r.operaciones);
}

function refreshAcabadoSummary() {
  const rows = Array.from(dom.acabadoTableBody?.querySelectorAll("tr") || []);
  let test = 0, tcot = 0;
  rows.forEach(row => {
    const data = readAcabadoRowValues(row);
    test += AppUtils.safeNumber(data.tiempoEstimado);
    tcot += AppUtils.safeNumber(data.tiempoCotizacion);
  });
  if (dom.summaryAcabado) {
    dom.summaryAcabado.estimated.textContent = AppUtils.formatNumber(test, 2);
    dom.summaryAcabado.cotizacion.textContent = AppUtils.formatNumber(tcot, 2);
  }
  refreshResumenSummary();
}

function renderAcabadosRows(rows) {
  appState.acabadoProtectionValue = resolveActiveProtectionValue(rows, APP_CONFIG.defaultProtection);
  if (dom.acabadoTableBody) dom.acabadoTableBody.innerHTML = "";
  if (rows.length) rows.forEach(r => appendAcabadoRow(r)); else appendAcabadoRow();
  refreshAcabadoSummary();
}

function buildInitialRows() {
  appState.costuraProtectionValue = APP_CONFIG.defaultProtection;
  appState.corteProtectionValue = APP_CONFIG.defaultProtection;
  appState.acabadoProtectionValue = APP_CONFIG.defaultProtection;
  dom.tableBody.innerHTML = "";
  if (dom.acabadoTableBody) dom.acabadoTableBody.innerHTML = "";

  for (let index = 0; index < APP_CONFIG.initialRows; index += 1) {
    appendRow();
  }
  
  buildInitialCorteRows();
  buildInitialAcabadoRows();
}

function createCosturaRow(rowData = null) {
  const fragment = dom.rowTemplate.content.cloneNode(true);
  const row = fragment.querySelector("tr");
  const codeInput = row.querySelector('[data-field="codigo"]');
  const protectionInput = row.querySelector('[data-field="proteccion"]');

  codeInput.addEventListener("change", handleCodeCommit);
  codeInput.addEventListener("blur", handleCodeCommit);
  codeInput.addEventListener("keydown", handleCodeKeydown);
  codeInput.addEventListener("contextmenu", handleCodeContextMenu);
  protectionInput.addEventListener("dblclick", handleProtectionDoubleClick);
  protectionInput.addEventListener("input", handleProtectionInput);
  protectionInput.addEventListener("blur", handleProtectionBlur);
  protectionInput.addEventListener("keydown", handleProtectionKeydown);

  writeRowValues(row, buildEmptyRow());
  if (rowData) {
    applyRowData(row, rowData);
  }

  syncRowInteractivity(row);
  return row;
}

function appendRow(rowData = null) {
  const row = createCosturaRow(rowData);
  dom.tableBody.appendChild(row);
  return row;
}

function insertRowAt(index, rowData = null) {
  const rows = getTableRows();
  const row = createCosturaRow(rowData);

  if (index >= rows.length) {
    dom.tableBody.appendChild(row);
    return row;
  }

  dom.tableBody.insertBefore(row, rows[index]);
  return row;
}

function focusRowCodeInput(row) {
  row?.querySelector('[data-field="codigo"]')?.focus();
}

function insertRowAbove(referenceRow) {
  if (!isEditableMode()) {
    return;
  }

  const rowIndex = getTableRows().indexOf(referenceRow);

  if (rowIndex === -1) {
    return;
  }

  const insertedRow = insertRowAt(rowIndex);
  ensureTrailingEmptyRow();
  refreshSummary();
  focusRowCodeInput(insertedRow);
}

function insertRowBelow(referenceRow) {
  if (!isEditableMode()) {
    return;
  }

  const rowIndex = getTableRows().indexOf(referenceRow);

  if (rowIndex === -1) {
    return;
  }

  const insertedRow = insertRowAt(rowIndex + 1);
  ensureTrailingEmptyRow();
  refreshSummary();
  focusRowCodeInput(insertedRow);
}

function deleteTableRow(rowToRemove) {
  if (!isEditableMode() || !rowToRemove) {
    return;
  }

  const rows = getTableRows();

  if (rows.length === 1) {
    writeRowValues(rows[0], buildEmptyRow());
    rows[0].classList.remove("is-invalid", "is-manual");
    refreshSummary();
    focusRowCodeInput(rows[0]);
    return;
  }

  const nextFocusRow = rowToRemove.nextElementSibling || rowToRemove.previousElementSibling;
  rowToRemove.remove();
  ensureTrailingEmptyRow();
  refreshSummary();
  focusRowCodeInput(nextFocusRow);
}

function handleCodeKeydown(event) {
  if (!isEditableMode()) {
    return;
  }
  if (event.key !== "Enter") {
    return;
  }

  event.preventDefault();
  handleCodeCommit(event);
  ensureTrailingEmptyRow();

  const rows = getTableRows();
  const currentRow = event.currentTarget.closest("tr");
  const currentIndex = rows.indexOf(currentRow);
  const nextRow = rows[currentIndex + 1];

  if (nextRow) {
    nextRow.querySelector('[data-field="codigo"]').focus();
  }
}

function handleCodeCommit(event) {
  if (!isEditableMode()) {
    return;
  }
  const row = event.currentTarget.closest("tr");
  const code = event.currentTarget.value.trim();

  if (!code) {
    writeRowValues(row, buildEmptyRow());
    row.querySelector('[data-field="codigo"]').value = "";
    row.classList.remove("is-invalid", "is-manual");
    refreshSummary();
    return;
  }

  const lookup = lookupOperation(code);

  if (!lookup) {
    row.classList.add("is-invalid");
    row.classList.remove("is-manual");
    writeRowValues(row, {
      ...buildEmptyRow(),
      codigo: code,
    });
    showToast(`No se encontrÃ³ el cÃ³digo ${code} en la hoja basedatos.`, "error");
    refreshSummary();
    return;
  }

  row.classList.remove("is-invalid");
  row.classList.toggle("is-manual", lookup.tipoPta === "*");
  writeRowValues(row, lookup);
  ensureTrailingEmptyRow();
  refreshSummary();
}

function lookupOperation(code) {
  const normalizedCode = AppUtils.normalizeKey(code);
  const source = appState.catalogByCode.get(normalizedCode);

  if (!source) {
    return null;
  }

  const tipoPta = appState.puntadasByTipoMaq.get(AppUtils.normalizeKey(source.tipoMaq)) || "";
  const tiempoEstimado = AppUtils.safeNumber(source.tiempoEstimado);
  const proteccion = appState.costuraProtectionValue;
  const isManual = tipoPta === "*";
  const tiempoMaq = isManual ? 0 : tiempoEstimado * proteccion;
  const tiempoManual = isManual ? tiempoEstimado * proteccion : 0;
  const tiempoCotizacion = tiempoMaq + tiempoManual;

  return {
    codigo: source.codigo,
    bloque: source.bloque,
    operaciones: source.operaciones,
    tiempoEstimado,
    tipoMaq: source.tipoMaq,
    proteccion,
    tipoPta,
    tiempoMaq,
    tiempoManual,
    tiempoCotizacion,
  };
}

function buildEmptyRow() {
  return {
    codigo: "",
    bloque: "",
    operaciones: "",
    tiempoEstimado: 0,
    tipoMaq: "",
    proteccion: appState.costuraProtectionValue,
    tipoPta: "",
    tiempoMaq: 0,
    tiempoManual: 0,
    tiempoCotizacion: 0,
  };
}

function applyRowData(row, rowData) {
  const normalizedRow = {
    codigo: rowData.codigo || "",
    bloque: rowData.bloque || "",
    operaciones: rowData.operaciones || "",
    tiempoEstimado: AppUtils.safeNumber(rowData.tiempoEstimado),
    tipoMaq: rowData.tipoMaq || "",
    proteccion: AppUtils.safeNumber(rowData.proteccion || APP_CONFIG.defaultProtection),
    tipoPta: rowData.tipoPta || "",
    tiempoMaq: AppUtils.safeNumber(rowData.tiempoMaq),
    tiempoManual: AppUtils.safeNumber(rowData.tiempoManual),
    tiempoCotizacion: AppUtils.safeNumber(rowData.tiempoCotizacion),
  };

  row.classList.toggle("is-manual", normalizedRow.tipoPta === "*");
  row.classList.remove("is-invalid");
  writeRowValues(row, normalizedRow);
}

function writeRowValues(row, rowData) {
  const fields = row.querySelectorAll("[data-field]");
  const isEmptyRow = !rowData.codigo && !rowData.operaciones;

  fields.forEach((field) => {
    const fieldName = field.dataset.field;
    const value = rowData[fieldName];

    if (
      [
        "tiempoEstimado",
        "proteccion",
        "tiempoMaq",
        "tiempoManual",
        "tiempoCotizacion",
      ].includes(fieldName)
    ) {
      if (isEmptyRow && fieldName !== "proteccion") {
        field.value = "";
        return;
      }

      const decimals =
        fieldName === "proteccion" ||
        fieldName === "tiempoEstimado" ||
        fieldName === "tiempoMaq" ||
        fieldName === "tiempoManual" ||
        fieldName === "tiempoCotizacion"
          ? 2
          : 3;
      field.value = AppUtils.formatNumber(value || 0, decimals);
      return;
    }

    field.value = value ?? "";
  });
}

function readRowValues(row) {
  const getValue = (fieldName) => row.querySelector(`[data-field="${fieldName}"]`).value;

  return {
    codigo: getValue("codigo").trim(),
    bloque: getValue("bloque").trim(),
    operaciones: getValue("operaciones").trim(),
    tiempoEstimado: AppUtils.safeNumber(getValue("tiempoEstimado")),
    tipoMaq: getValue("tipoMaq").trim(),
    proteccion: AppUtils.safeNumber(getValue("proteccion")) || APP_CONFIG.defaultProtection,
    tipoPta: getValue("tipoPta").trim(),
    tiempoMaq: AppUtils.safeNumber(getValue("tiempoMaq")),
    tiempoManual: AppUtils.safeNumber(getValue("tiempoManual")),
    tiempoCotizacion: AppUtils.safeNumber(getValue("tiempoCotizacion")),
  };
}

function getTableRows() {
  return Array.from(dom.tableBody.querySelectorAll("tr"));
}

function getFilledRows() {
  return getTableRows()
    .map(readRowValues)
    .filter((row) => row.codigo || row.operaciones);
}

function ensureTrailingEmptyRow() {
  if (!isEditableMode()) {
    return;
  }

  const rows = getTableRows();
  const hasEmptyRow = rows.some((row) => !row.querySelector('[data-field="codigo"]').value.trim());

  if (!hasEmptyRow) {
    appendRow();
  }
}

function isEditableMode() {
  return EDITABLE_MODES.has(appState.interactionMode);
}

function getActiveRecord() {
  return appState.searchResults.find((item) => item.key === appState.activeRecordKey) || null;
}

function getSaveMode() {
  if (appState.interactionMode === "search_edit") {
    return "update_selected";
  }

  if (appState.interactionMode === "search_new_version") {
    return "upsert_proto_version";
  }

  return "create";
}

function buildRecordLocator() {
  const record = getActiveRecord();

  if (!record) {
    return null;
  }

  return {
    recordId: record.recordId || "",
    rowNumber: record.rowNumber || "",
    proto: record.proto || "",
    version: record.version || "",
  };
}

function syncRowInteractivity(row) {
  const codeInput = row.querySelector('[data-field="codigo"]');
  const protectionInput = row.querySelector('[data-field="proteccion"]');

  if (!codeInput) {
    return;
  }

  codeInput.readOnly = !isEditableMode();
  codeInput.tabIndex = isEditableMode() ? 0 : -1;

  if (protectionInput) {
    lockProtectionEditing(protectionInput);
  }
}

function applyTableInteractivity() {
  const canEdit = isEditableMode();

  getTableRows().forEach(syncRowInteractivity);
  dom.addRowBtn.disabled = !canEdit;
  dom.removeRowBtn.disabled = !canEdit;
  dom.clearTableBtn.disabled = !canEdit;
  dom.addCorteRowBtn.disabled = !canEdit;
  dom.removeCorteRowBtn.disabled = !canEdit;
  dom.clearCorteTableBtn.disabled = !canEdit;
  dom.addAcabadoRowBtn.disabled = !canEdit;
  dom.removeAcabadoRowBtn.disabled = !canEdit;
  dom.clearAcabadoTableBtn.disabled = !canEdit;
}

function setInteractionMode(mode) {
  appState.interactionMode = mode;
  closeFloatingMenus();

  const canEdit = isEditableMode();
  const allowProtoEdit = mode === "create";
  const allowVersionEdit = mode === "create" || mode === "search_new_version";

  dom.form.cliente.disabled = !canEdit;
  dom.form.proto.readOnly = !allowProtoEdit;
  dom.form.version.readOnly = !allowVersionEdit;
  dom.form.idem.readOnly = !canEdit;
  dom.form.descripcion.readOnly = !canEdit;
  dom.form.estilo.readOnly = !canEdit;
  dom.form.tela.readOnly = !canEdit;
  dom.form.realizadoPor.readOnly = !canEdit;
  dom.form.produccionEstimada.readOnly = !canEdit;
  dom.form.rutasProcesos.readOnly = !canEdit;
  dom.form.clienteCustom.readOnly = !canEdit;
  dom.form.clienteCustom.disabled = !canEdit;
  dom.form.clienteDelete.disabled = !canEdit;

  if (!canEdit) {
    hideCustomClientInput();
  }

  dom.saveBtn.classList.toggle("hidden", !canEdit);

  const saveTitle =
    mode === "search_edit"
      ? "Actualizar versiÃ³n seleccionada"
      : mode === "search_new_version"
        ? "Guardar nueva versiÃ³n"
        : "Guardar cotizaciÃ³n";

  dom.saveBtn.setAttribute(
    "title",
    mode === "search_edit"
      ? "Actualizar version seleccionada"
      : mode === "search_new_version"
        ? "Guardar nueva version"
        : "Guardar cotizacion"
  );
  dom.saveBtn.setAttribute("aria-label", dom.saveBtn.getAttribute("title"));
  dom.formPanel.classList.toggle("is-readonly", !canEdit);
  dom.costuraPanel.classList.toggle("is-readonly", !canEdit);
  dom.cortePanel.classList.toggle("is-readonly", !canEdit);
  dom.acabadoPanel.classList.toggle("is-readonly", !canEdit);
  dom.resumenPanel.classList.toggle("is-readonly", !canEdit);

  applyTableInteractivity();
  updatePrototypeClearButtonVisibility();
  updateDeleteClientButton();
  updateVersionActionsState();
  updateVersionInfo();
}

function renderCosturaRows(rows, includeTrailingEmpty = false) {
  closeFloatingMenus();
  appState.costuraProtectionValue = resolveActiveProtectionValue(rows, APP_CONFIG.defaultProtection);
  dom.tableBody.innerHTML = "";

  if (rows.length) {
    rows.forEach((row) => appendRow(row));
  } else {
    appendRow();
  }

  if (includeTrailingEmpty) {
    ensureTrailingEmptyRow();
  }

  applyTableInteractivity();
  refreshSummary();
}

function resolveRecordKey(records, options = {}) {
  if (options.preferredRecordId) {
    const preferredRecord = records.find((record) => record.recordId === options.preferredRecordId);

    if (preferredRecord) {
      return preferredRecord.key;
    }
  }

  if (options.preferredVersion) {
    const normalizedVersion = AppUtils.normalizeKey(options.preferredVersion);
    const preferredRecord = records.find(
      (record) => AppUtils.normalizeKey(record.version || "") === normalizedVersion
    );

    if (preferredRecord) {
      return preferredRecord.key;
    }
  }

  return records[records.length - 1].key;
}

function updateVersionActionsState() {
  const hasActiveRecord = Boolean(getActiveRecord());

  dom.editVersionBtn.disabled = !hasActiveRecord || appState.interactionMode === "search_edit";
  dom.newVersionBtn.disabled = !hasActiveRecord || appState.interactionMode === "search_new_version";
  dom.printRecordBtn.disabled = !hasActiveRecord;
}

function getVersionModeLabel(mode = appState.interactionMode) {
  if (mode === "search_edit") {
    return "Editando";
  }

  if (mode === "search_new_version" || mode === "create") {
    return "Nueva Version";
  }

  return "Lectura";
}

function getVersionSortValue(record) {
  return String(record && record.version ? record.version : "").trim();
}

function getSavedAtTimestamp(record) {
  const date = new Date(record && record.savedAt ? record.savedAt : "");
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function sortVersionRecords(records) {
  return [...records].sort((left, right) => {
    const leftVersion = getVersionSortValue(left);
    const rightVersion = getVersionSortValue(right);

    if (leftVersion && rightVersion) {
      const versionCompare = VERSION_COLLATOR.compare(leftVersion, rightVersion);
      if (versionCompare !== 0) {
        return versionCompare;
      }
    } else if (leftVersion) {
      return -1;
    } else if (rightVersion) {
      return 1;
    }

    const savedAtCompare = getSavedAtTimestamp(left) - getSavedAtTimestamp(right);
    if (savedAtCompare !== 0) {
      return savedAtCompare;
    }

    return VERSION_COLLATOR.compare(left.recordId || "", right.recordId || "");
  });
}

function updateVersionInfo() {
  if (!appState.searchResults.length) {
    dom.versionInfo.textContent = "";
    dom.versionSubinfo.textContent = "";
    return;
  }

  dom.versionInfo.textContent = "";
  dom.versionSubinfo.textContent = getVersionModeLabel();
}

function updateFormRecordInfo(record) {
  if (!dom.formRecordInfo) {
    return;
  }

  const meta = AppUtils.formatVersionMeta(record);
  dom.formRecordInfo.textContent = meta === "Sin fecha" ? "" : `FECHA REGISTRO: ${meta}`;
}

function clearDisplayedRecord() {
  setClientValue("");
  dom.form.proto.value = "";
  dom.form.version.value = "";
  dom.form.idem.value = "";
  dom.form.descripcion.value = "";
  dom.form.estilo.value = "";
  dom.form.tela.value = "";
  dom.form.realizadoPor.value = "";
  dom.form.produccionEstimada.value = "";
  dom.form.rutasProcesos.value = "";
  updateFormRecordInfo(null);
  renderCosturaRows([], false);
  renderCorteRows([]);
  renderAcabadosRows([]);
}

function refreshSummary() {
  const rows = getFilledRows();

  const totals = rows.reduce(
    (accumulator, row) => {
      accumulator.estimated += AppUtils.safeNumber(row.tiempoEstimado);
      accumulator.maq += AppUtils.safeNumber(row.tiempoMaq);
      accumulator.manual += AppUtils.safeNumber(row.tiempoManual);
      accumulator.cotizacion += AppUtils.safeNumber(row.tiempoCotizacion);
      return accumulator;
    },
    {
      estimated: 0,
      maq: 0,
      manual: 0,
      cotizacion: 0,
    }
  );

  dom.summary.estimatedFooter.textContent = AppUtils.formatNumber(totals.estimated, 2);
  dom.summary.maqFooter.textContent = AppUtils.formatNumber(totals.maq, 2);
  dom.summary.manualFooter.textContent = AppUtils.formatNumber(totals.manual, 2);
  dom.summary.cotizacionFooter.textContent = AppUtils.formatNumber(totals.cotizacion, 2);
  refreshResumenSummary();
}

function refreshResumenSummary() {
  if (!dom.summaryResumen) {
    return;
  }

  const corteTotal = AppUtils.safeNumber(dom.summaryCorte?.cotizacion?.textContent);
  const costuraMaq = AppUtils.safeNumber(dom.summary?.maqFooter?.textContent);
  const costuraManual = AppUtils.safeNumber(dom.summary?.manualFooter?.textContent);
  const costuraTotal = AppUtils.safeNumber(dom.summary?.cotizacionFooter?.textContent);
  const acabadosTotal = AppUtils.safeNumber(dom.summaryAcabado?.cotizacion?.textContent);
  const granTotal = corteTotal + costuraTotal + acabadosTotal;

  dom.summaryResumen.corte.textContent = AppUtils.formatNumber(corteTotal, 2);
  dom.summaryResumen.costuraMaq.textContent = AppUtils.formatNumber(costuraMaq, 2);
  dom.summaryResumen.costuraManual.textContent = AppUtils.formatNumber(costuraManual, 2);
  dom.summaryResumen.costuraTotal.textContent = AppUtils.formatNumber(costuraTotal, 2);
  dom.summaryResumen.acabados.textContent = AppUtils.formatNumber(acabadosTotal, 2);
  dom.summaryResumen.total.textContent = AppUtils.formatNumber(granTotal, 2);
}

function handleNewPrototype() {
  if (appState.interactionMode === "create") {
    switchView("editor");
    return;
  }

  resetPrototypeEditor();
  setInteractionMode("create");
  switchView("editor");
  dom.form.proto.focus();
}

function resetPrototypeEditor() {
  resetForm();
  buildInitialRows();
  refreshSummary();
  appState.searchResults = [];
  appState.activeRecordKey = null;
  appState.lastSearchProto = "";
  dom.versionTabs.innerHTML = "";
  dom.versionInfo.textContent = "";
  dom.versionSubinfo.textContent = "";
}

function resetForm() {
  clearDisplayedRecord();
  dom.searchSummary.textContent = "Busca un PROTO para listar sus versiones guardadas.";
}

function updatePrototypeClearButtonVisibility() {
  if (!dom.clearPrototypeBtn && !dom.copyPrototypeBtn) {
    return;
  }

  const shouldShow = appState.currentView === "editor" && appState.interactionMode === "create";
  [dom.copyPrototypeBtn, dom.clearPrototypeBtn].forEach((button) => {
    if (!button) {
      return;
    }

    button.classList.toggle("hidden", !shouldShow);
    button.disabled = !shouldShow;
  });
}

function handleClearPrototypeWithConfirm() {
  if (appState.currentView !== "editor" || appState.interactionMode !== "create") {
    return;
  }

  openConfirmModal("Se limpiaran todos los datos ingresados en Nuevo prototipo. Deseas continuar?", () => {
    resetPrototypeEditor();
    switchView("editor");
    dom.form.proto.focus();
  }, {
    title: "Limpieza de datos",
    acceptLabel: "Limpiar",
  });
}

function openCopyPrototypeModal(initialProto = "") {
  if (appState.currentView !== "editor" || appState.interactionMode !== "create" || !dom.copyPrototypeModal) {
    return;
  }

  closeConfirmModal();
  dom.copyPrototypeInput.value = normalizeProtoValue(initialProto);
  dom.copyPrototypeModal.classList.remove("hidden");
  dom.copyPrototypeModal.setAttribute("aria-hidden", "false");
  dom.copyPrototypeInput.focus();
  dom.copyPrototypeInput.select();
}

function closeCopyPrototypeModal() {
  if (!dom.copyPrototypeModal) {
    return;
  }

  dom.copyPrototypeModal.classList.add("hidden");
  dom.copyPrototypeModal.setAttribute("aria-hidden", "true");
}

function handleCopyPrototypeInputKeydown(event) {
  if (event.key !== "Enter") {
    return;
  }

  event.preventDefault();
  handleCopyPrototypeAccept();
}

async function handleCopyPrototypeAccept() {
  const sourceProto = normalizeProtoValue(dom.copyPrototypeInput.value.trim());
  dom.copyPrototypeInput.value = sourceProto;

  if (!/^\d{5}$/.test(sourceProto)) {
    showToast("Ingresa un PROTO valido de 5 digitos para copiar.", "error");
    dom.copyPrototypeInput.focus();
    dom.copyPrototypeInput.select();
    return;
  }

  closeCopyPrototypeModal();
  showLoadingModal(`Copiando PROTO ${sourceProto}`);

  try {
    const records = await fetchProtoHistoryRecords(sourceProto);

    if (!records.length) {
      throw new Error(`No se encontraron versiones registradas para el PROTO ${sourceProto}.`);
    }

    const sourceRecord = records[records.length - 1];
    populateNewPrototypeFromReference(sourceRecord, sourceProto);
    showToast(
      `Se copiaron los datos del PROTO ${sourceProto} usando la version ${sourceRecord.version || "mas reciente"}.`,
      "success"
    );
  } catch (error) {
    openCopyPrototypeModal(sourceProto);
    showToast(error.message || "No se pudo copiar el PROTO indicado.", "error");
  } finally {
    hideLoadingModal();
  }
}

function populateNewPrototypeFromReference(record, sourceProto) {
  resetPrototypeEditor();
  setClientValue(record.cliente || "");
  dom.form.proto.value = "";
  dom.form.version.value = "";
  dom.form.idem.value = normalizeIdemValue(sourceProto);
  setUppercaseInputValue(dom.form.descripcion, record.descripcion || "");
  setUppercaseInputValue(dom.form.estilo, record.estilo || "");
  setUppercaseInputValue(dom.form.tela, record.tela || "");
  setUppercaseInputValue(dom.form.realizadoPor, record.realizadoPor || "");
  dom.form.produccionEstimada.value = formatProductionEstimadaValue(record.produccionEstimada || "");
  setUppercaseInputValue(dom.form.rutasProcesos, record.rutasProcesos || "");
  updateFormRecordInfo(null);
  renderCosturaRows(record.costuraRows || [], true);
  renderCorteRows(record.corteRows || []);
  renderAcabadosRows(record.acabadoRows || []);
  setInteractionMode("create");
  switchView("editor");
  dom.form.proto.focus();
}

function switchView(view, options = {}) {
  const isSearch = view === "search";
  const isMaster = view === "master";
  const focusSearch = options.focusSearch !== false;

  appState.currentView = view;
  closeFloatingMenus();
  dom.searchVersionGrid.classList.toggle("hidden", !isSearch);
  dom.searchPanel.classList.toggle("hidden", !isSearch);
  dom.versionPanel.classList.toggle("hidden", !isSearch);
  dom.formPanel.classList.toggle("hidden", isMaster);
  
  if (dom.panelTabs) dom.panelTabs.classList.toggle("hidden", isMaster);
  
  if (isMaster) {
    dom.costuraPanel.classList.add("hidden");
    dom.cortePanel.classList.add("hidden");
    dom.acabadoPanel.classList.add("hidden");
    dom.resumenPanel.classList.add("hidden");
  } else {
    const activeTab = dom.panelTabs ? dom.panelTabs.querySelector('.is-active') : null;
    const targetId = activeTab ? activeTab.dataset.target : 'cortePanel';
    dom.costuraPanel.classList.toggle("hidden", targetId !== 'costuraPanel');
    dom.cortePanel.classList.toggle("hidden", targetId !== 'cortePanel');
    dom.acabadoPanel.classList.toggle("hidden", targetId !== 'acabadoPanel');
    dom.resumenPanel.classList.toggle("hidden", targetId !== 'resumenPanel');
  }

  dom.masterPanel.classList.toggle("hidden", !isMaster);
  dom.newPrototypeBtn.classList.toggle("is-active", view === "editor");
  dom.openSearchBtn.classList.toggle("is-active", isSearch);
  dom.openMasterBtn.classList.toggle("is-active", isMaster);

  if (isSearch && focusSearch) {
    dom.searchProtoInput.focus();
  }

  updatePrototypeClearButtonVisibility();
}

async function handleSearch() {
  if (!AppsScriptAPI.isConfigured()) {
    showToast("Primero configura WEB_APP_URL para activar la bÃºsqueda.", "error");
    return;
  }

  const proto = normalizeSearchProto(dom.searchProtoInput.value.trim() || dom.form.proto.value.trim());
  dom.searchProtoInput.value = proto;

  if (!proto) {
    showToast("Ingresa un PROTO para buscar.", "error");
    return;
  }

  if (!/^\d{5}$/.test(proto)) {
    showToast("El PROTO debe ser un numero entero de 5 digitos.", "error");
    dom.searchProtoInput.focus();
    dom.searchProtoInput.select();
    return;
  }

  dom.searchBtn.disabled = true;
  dom.searchSummary.textContent = `Buscando historial para ${proto}...`;

  try {
    await loadProtoRecords(proto, { mode: "search_readonly" });
    showToast(`Historial cargado para ${proto}.`, "success");
  } catch (error) {
    dom.searchSummary.textContent = "OcurriÃ³ un error al consultar Google Sheets.";
    showToast(error.message || "No se pudo completar la bÃºsqueda.", "error");
  } finally {
    dom.searchBtn.disabled = false;
  }
}

function buildProtoHistoryRecords(records = []) {
  return sortVersionRecords(
    (records || []).map((record, index) => ({
      ...record,
      key: record.recordId || `${record.version || "SIN_VERSION"}-${index}`,
      costuraRows:
        Array.isArray(record.costuraRows) && record.costuraRows.length
          ? record.costuraRows
          : AppUtils.parseCosturaCsv ? AppUtils.parseCosturaCsv(record.costuraCsv || "") : [],
      corteRows:
        Array.isArray(record.corteRows) && record.corteRows.length
          ? record.corteRows
          : AppUtils.parseCorteCsv ? AppUtils.parseCorteCsv(record.corteCsv || "") : [],
      acabadoRows:
        Array.isArray(record.acabadoRows) && record.acabadoRows.length
          ? record.acabadoRows
          : AppUtils.parseAcabadoCsv ? AppUtils.parseAcabadoCsv(record.acabadoCsv || "") : [],
    }))
  );
}

async function fetchProtoHistoryRecords(proto) {
  const response = await AppsScriptAPI.searchByProto(String(proto || "").trim());
  const payload = response.data || response;

  if (!response.success && !Array.isArray(payload.records)) {
    throw new Error(response.message || "No se pudo recuperar el historial.");
  }

  return buildProtoHistoryRecords(payload.records || []);
}

async function loadProtoRecords(proto, options = {}) {
  const cleanProto = String(proto || "").trim();
  const records = await fetchProtoHistoryRecords(cleanProto);

  appState.searchResults = records;
  appState.lastSearchProto = cleanProto;
  dom.searchProtoInput.value = cleanProto;
  switchView("search", { focusSearch: false });

   if (!records.length) {
     appState.activeRecordKey = null;
     dom.versionTabs.innerHTML = "";
     dom.versionInfo.textContent = "";
     clearDisplayedRecord();
     setInteractionMode("search_readonly");
     dom.searchSummary.textContent = `No se encontraron versiones registradas para el PROTO ${cleanProto}.`;
     return [];
   }

   const targetKey = resolveRecordKey(records, options);
   loadRecord(targetKey, options.mode || "search_readonly");
   dom.searchSummary.textContent = `Se encontraron ${records.length} versiÃ³n(es) para el PROTO ${cleanProto}.`;

   return records;
}

function renderVersionTabs(records) {
  dom.versionTabs.innerHTML = "";

  records.forEach((record) => {
    const statusLabel = record.key === appState.activeRecordKey ? getVersionModeLabel() : "Lectura";
    const versionLabel = String(record.version || "").trim() || "Sin versiÃ³n";
    const metaLabel = AppUtils.formatVersionMeta(record);
    const button = document.createElement("button");
    const pill = document.createElement("span");

    button.type = "button";
    button.className = "version-tab";
    button.dataset.key = record.key;
    button.setAttribute(
      "aria-label",
      metaLabel === "Sin fecha"
        ? `VersiÃ³n ${versionLabel}. ${statusLabel}.`
        : `VersiÃ³n ${versionLabel}. ${statusLabel}. Fecha de registro ${metaLabel}.`
    );

    pill.className = "version-pill";
    pill.textContent = versionLabel;

    button.appendChild(pill);
    button.addEventListener("click", () => loadRecord(record.key, "search_readonly"));
    dom.versionTabs.appendChild(button);
  });

  updateVersionActionsState();
  updateVersionInfo();
}

function loadRecord(recordKey, mode = "search_readonly") {
  const record = appState.searchResults.find((item) => item.key === recordKey);

  if (!record) {
    return;
  }

  appState.activeRecordKey = recordKey;

  setClientValue(record.cliente || "");
  dom.form.proto.value = normalizeProtoValue(record.proto || "");
  dom.form.version.value = normalizeShortIntegerValue(record.version || "");
  dom.form.idem.value = normalizeIdemValue(record.idem || "");
  setUppercaseInputValue(dom.form.descripcion, record.descripcion || "");
  setUppercaseInputValue(dom.form.estilo, record.estilo || "");
  setUppercaseInputValue(dom.form.tela, record.tela || "");
  setUppercaseInputValue(dom.form.realizadoPor, record.realizadoPor || "");
  dom.form.produccionEstimada.value = formatProductionEstimadaValue(record.produccionEstimada || "");
  setUppercaseInputValue(dom.form.rutasProcesos, record.rutasProcesos || "");

  updateFormRecordInfo(record);
  renderCosturaRows(record.costuraRows || [], mode !== "search_readonly");
  renderCorteRows(record.corteRows || [], { hideZeroValues: mode === "search_readonly" });
  renderAcabadosRows(record.acabadoRows || []);

  if (mode === "search_new_version") {
    dom.form.version.value = "";
  }

   setInteractionMode(mode);
   // Re-render version tabs to update status labels based on interaction mode
   renderVersionTabs(appState.searchResults);
   syncActiveVersionTab();
}

function handleEditVersion() {
  const record = getActiveRecord();

  if (!record) {
    return;
  }

  loadRecord(record.key, "search_edit");
  dom.form.descripcion.focus();
}

function handleNewVersionFromRecord() {
  const record = getActiveRecord();

  if (!record) {
    return;
  }

  loadRecord(record.key, "search_new_version");
  dom.form.version.focus();
  dom.form.version.select();
  showToast("Ingresa la nueva versiÃ³n y guarda para crearla o sobrescribirla.", "info");
}

function syncActiveVersionTab() {
  dom.versionTabs.querySelectorAll(".version-tab").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.key === appState.activeRecordKey);
  });

  updateVersionActionsState();
}

function clearTableWithConfirm() {
  if (!isEditableMode()) {
    return;
  }

  const hasData = getFilledRows().length > 0;

  if (!hasData) {
    buildInitialRows();
    refreshSummary();
    applyTableInteractivity();
    return;
  }

  openConfirmModal("Esta seguro que quiere eliminar todos los datos de la tabla?", () => {
    buildInitialRows();
    refreshSummary();
    applyTableInteractivity();
  }, {
    acceptLabel: "Eliminar",
  });
}

function removeLastRow() {
  if (!isEditableMode()) {
    return;
  }
  const rows = getTableRows();

  if (!rows.length) {
    appendRow();
    refreshSummary();
    return;
  }

  if (rows.length === 1) {
    writeRowValues(rows[0], buildEmptyRow());
    rows[0].classList.remove("is-invalid", "is-manual");
    refreshSummary();
    return;
  }

  const rowToRemove = rows[rows.length - 1];
  const rowData = readRowValues(rowToRemove);
  const removedHadData = Boolean(rowData.codigo || rowData.operaciones);

  rowToRemove.remove();

  if (removedHadData) {
    ensureTrailingEmptyRow();
  }

  refreshSummary();
}

function openConfirmModal(message, onAccept, options = {}) {
  const {
    title = "Confirmación",
    acceptLabel = "Aceptar",
    cancelLabel = "Cancelar",
    hideCancel = false,
  } = options;

  confirmAction = typeof onAccept === "function" ? onAccept : null;
  dom.confirmModalTitle.textContent = title;
  dom.confirmModalMessage.textContent = message;
  dom.confirmModalCancelBtn.textContent = cancelLabel;
  dom.confirmModalAcceptBtn.textContent = acceptLabel;
  dom.confirmModalCancelBtn.classList.toggle("hidden", hideCancel);
  dom.confirmModal.classList.remove("hidden");
  dom.confirmModal.setAttribute("aria-hidden", "false");
  dom.confirmModalAcceptBtn.focus();
}

function openAlertModal(message, onAccept = null) {
  openConfirmModal(message, onAccept, {
    title: "Completar todos los campos",
    acceptLabel: "Aceptar",
    hideCancel: true,
  });
}

function closeConfirmModal() {
  confirmAction = null;
  dom.confirmModal.classList.add("hidden");
  dom.confirmModal.setAttribute("aria-hidden", "true");
  dom.confirmModalTitle.textContent = "Confirmación";
  dom.confirmModalMessage.textContent = "Esta seguro que quiere eliminar todos los datos de la tabla?";
  dom.confirmModalCancelBtn.textContent = "Cancelar";
  dom.confirmModalAcceptBtn.textContent = "Aceptar";
  dom.confirmModalCancelBtn.classList.remove("hidden");
}

function handleConfirmAccept() {
  const action = confirmAction;
  closeConfirmModal();

  if (typeof action === "function") {
    action();
  }
}

function focusFormControl(element) {
  if (!element) {
    return;
  }

  if (typeof element.focus === "function") {
    element.focus();
  }

  if (typeof element.select === "function" && !element.readOnly && !element.disabled) {
    element.select();
  }
}

function getFirstIncompleteRequiredFormField(formData) {
  const requiredFields = [
    {
      value: formData.cliente,
      element: dom.form.cliente.value === NEW_CLIENT_OPTION_VALUE ? dom.form.clienteCustom : dom.form.cliente,
    },
    { value: formData.proto, element: dom.form.proto },
    { value: formData.version, element: dom.form.version },
    { value: formData.idem, element: dom.form.idem },
    { value: formData.estilo, element: dom.form.estilo },
    { value: formData.tela, element: dom.form.tela },
    { value: formData.descripcion, element: dom.form.descripcion },
    { value: formData.realizadoPor, element: dom.form.realizadoPor },
    { value: formData.produccionEstimada, element: dom.form.produccionEstimada },
    { value: formData.rutasProcesos, element: dom.form.rutasProcesos },
  ];

  return requiredFields.find((field) => !String(field.value || "").trim()) || null;
}

function isCosturaRowEmpty(rowElement) {
  return ["codigo", "bloque", "operaciones"].every(
    (fieldName) => !rowElement.querySelector(`[data-field="${fieldName}"]`).value.trim()
  );
}

function isCorteRowEmpty(rowElement) {
  return ["operaciones", "tiempoEstimadoCorte", "tiempoEstimadoHabilitado", "area"].every(
    (fieldName) => !rowElement.querySelector(`[data-field="${fieldName}"]`).value.trim()
  );
}

function isAcabadoRowEmpty(rowElement) {
  return ["operaciones", "tiempoEstimado"].every(
    (fieldName) => !rowElement.querySelector(`[data-field="${fieldName}"]`).value.trim()
  );
}

function isCosturaRowComplete(rowElement) {
  const rowData = readRowValues(rowElement);
  const hasRequiredText = [rowData.codigo, rowData.bloque, rowData.operaciones, rowData.tipoMaq, rowData.tipoPta].every(
    (value) => String(value || "").trim()
  );
  const hasNumericValues = ["tiempoEstimado", "proteccion", "tiempoMaq", "tiempoManual", "tiempoCotizacion"].every(
    (fieldName) => Number.isFinite(rowData[fieldName])
  );
  const isComplete = !rowElement.classList.contains("is-invalid") && hasRequiredText && hasNumericValues;

  if (!isCosturaRowEmpty(rowElement)) {
    rowElement.classList.toggle("is-invalid", !isComplete);
  }

  return isComplete;
}

function validateSaveRequirements(formData) {
  const firstIncompleteField = getFirstIncompleteRequiredFormField(formData);

  if (firstIncompleteField) {
    return {
      isValid: false,
      focusElement: firstIncompleteField.element,
      message:
        "Para guardar la cotización debes completar todos los campos de Información general de la prenda y al menos una fila completa en Costura.",
    };
  }

  const formatRules = [
    {
      value: formData.proto,
      pattern: /^\d{1,2}$/,
      element: dom.form.proto,
      message: "El campo PROTO debe ser un número entero de 5 dígitos.",
    },
    {
      value: formData.version,
      pattern: /^\d{1,2}$/,
      element: dom.form.version,
      message: "El campo Versión debe ser un número entero de 1 o 2 dígitos.",
    },
    {
      value: formData.idem,
      pattern: /^\d{5}$/,
      element: dom.form.idem,
      message: "El campo IDEM debe ser un numero entero de 5 digitos.",
    },
    {
      value: formData.produccionEstimada,
      pattern: /^\d{1,3}(,\d{3})*$/,
      element: dom.form.produccionEstimada,
      message: "Producción estimada debe ser un número entero y mostrarse en miles con coma.",
    },
  ];

  const firstInvalidField = formatRules.find(
    (rule) => !rule.pattern.test(String(rule.value || "").trim())
  );

  if (firstInvalidField) {
    return {
      isValid: false,
      focusElement: firstInvalidField.element,
      message: firstInvalidField.message,
    };
  }

  const nonEmptyRows = getTableRows().filter((rowElement) => !isCosturaRowEmpty(rowElement));
  const hasCorteRows = getCorteFilledRows().length > 0;
  const hasAcabadoRows = getAcabadoFilledRows().length > 0;

  if (!nonEmptyRows.length && !hasCorteRows && !hasAcabadoRows) {
    return {
      isValid: false,
      focusElement: getTableRows()[0]?.querySelector('[data-field="codigo"]') || null,
      message:
        "Para guardar la cotización debes completar todos los campos de Información general y al menos una fila completa en cualquier tabla.",
    };
  }

  const rowStates = nonEmptyRows.map((rowElement) => ({
    rowElement,
    isComplete: isCosturaRowComplete(rowElement),
  }));
  const firstIncompleteRow = rowStates.find((item) => !item.isComplete);
  const hasCompleteRow = rowStates.some((item) => item.isComplete);

  if (!hasCompleteRow || firstIncompleteRow) {
    return {
      isValid: false,
      focusElement: (firstIncompleteRow ? firstIncompleteRow.rowElement : nonEmptyRows[0]).querySelector('[data-field="codigo"]') || null,
      message:
        "Para guardar la cotización debes completar todos los campos de Información general de la prenda y al menos una fila completa en Costura.",
    };
  }

  return {
    isValid: true,
    focusElement: null,
    message: "",
  };
}

async function handleSave() {
  if (!AppsScriptAPI.isConfigured()) {
    showToast("Actualiza WEB_APP_URL antes de guardar en Google Sheets.", "error");
    return;
  }

  if (!isEditableMode()) {
    showToast("Usa Editar o Nueva versiÃ³n para habilitar el guardado.", "info");
    return;
  }

  commitCustomClient();
  normalizeGeneralFormFields();
  const formData = readFormData();
  const validation = validateSaveRequirements(formData);

  if (!validation.isValid) {
    openAlertModal(
      validation.message ||
        "Para guardar la cotización debes completar todos los campos de Información general de la prenda y al menos una fila completa en Costura.",
      () => {
        focusFormControl(validation.focusElement);
      }
    );
    return;
  }

  const rows = getFilledRows();
  const corteRows = getCorteFilledRows();
  const acabadoRows = getAcabadoFilledRows();
  const saveMode = getSaveMode();

  if (saveMode === "upsert_proto_version" && !formData.version) {
    showToast("El campo VersiÃ³n es obligatorio para guardar una nueva versiÃ³n.", "error");
    dom.form.version.focus();
    return;
  }

  if (!rows.length && !corteRows.length && !acabadoRows.length) {
    showToast("Ingresa al menos una operaciÃ³n en alguna de las tablas.", "error");
    return;
  }

  if (rows.some((row) => !row.codigo || !row.operaciones) ||
      corteRows.some((row) => !row.operaciones) ||
      acabadoRows.some((row) => !row.operaciones)) {
    showToast("Hay filas incompletas. Revísalas antes de guardar.", "error");
    return;
  }

  const payload = {
    saveMode,
    recordLocator: buildRecordLocator(),
    form: formData,
    costuraRows: rows,
    corteRows: corteRows,
    acabadoRows: acabadoRows,
    costuraCsv: AppUtils.serializeCosturaRows(rows),
    corteCsv: AppUtils.serializeCorteRows ? AppUtils.serializeCorteRows(corteRows) : null,
    acabadoCsv: AppUtils.serializeAcabadoRows ? AppUtils.serializeAcabadoRows(acabadoRows) : null,
    summary: {
      totalOperaciones: rows.length + corteRows.length + acabadoRows.length,
      totalTiempoCotizacion: rows.reduce((total, row) => total + AppUtils.safeNumber(row.tiempoCotizacion), 0) +
                             corteRows.reduce((total, row) => total + AppUtils.safeNumber(row.tiempoCotizacion), 0) +
                             acabadoRows.reduce((total, row) => total + AppUtils.safeNumber(row.tiempoCotizacion), 0),
    },
  };

  dom.saveBtn.disabled = true;
  dom.saveBtn.classList.add("is-saving");
  dom.saveBtn.setAttribute("aria-busy", "true");
  dom.saveBtn.setAttribute("title", "Guardando...");

  try {
    const response = await AppsScriptAPI.saveCotizacion(payload);

    if (!response.success) {
      throw new Error(response.message || "No se pudo guardar la cotizaciÃ³n.");
    }

    if (saveMode !== "create") {
      await loadProtoRecords(formData.proto, {
        preferredRecordId: response.recordId,
        preferredVersion: formData.version,
        mode: "search_readonly",
      });
    } else {
      resetPrototypeEditor();
      setInteractionMode("create");
      switchView("editor");
      dom.form.proto.focus();
    }

    showToast(response.message || "CotizaciÃ³n guardada correctamente.", "success");
  } catch (error) {
    showToast(error.message || "OcurriÃ³ un error al guardar la cotizaciÃ³n.", "error");
  } finally {
    dom.saveBtn.disabled = false;
    dom.saveBtn.classList.remove("is-saving");
    dom.saveBtn.removeAttribute("aria-busy");
    dom.saveBtn.setAttribute("title", dom.saveBtn.getAttribute("aria-label") || "Guardar cotizaciÃ³n");
  }
}

function readFormData() {
  const clienteValue =
    dom.form.cliente.value === NEW_CLIENT_OPTION_VALUE
      ? dom.form.clienteCustom.value.trim()
      : dom.form.cliente.value.trim();
  const normalizedClienteValue = normalizeUppercaseTextValue(clienteValue).trim();

  return {
    cliente: normalizedClienteValue,
    proto: normalizeProtoValue(dom.form.proto.value),
    version: normalizeShortIntegerValue(dom.form.version.value),
    idem: normalizeIdemValue(dom.form.idem.value),
    descripcion: normalizeUppercaseTextValue(dom.form.descripcion.value).trim(),
    estilo: normalizeUppercaseTextValue(dom.form.estilo.value).trim(),
    tela: normalizeUppercaseTextValue(dom.form.tela.value).trim(),
    realizadoPor: normalizeUppercaseTextValue(dom.form.realizadoPor.value).trim(),
    produccionEstimada: formatProductionEstimadaValue(dom.form.produccionEstimada.value),
    rutasProcesos: normalizeUppercaseTextValue(dom.form.rutasProcesos.value).trim(),
  };
}

function clearPrintRecordMode() {
  document.body.classList.remove("print-record-mode");
}

function handlePrintRecord() {
  const record = getActiveRecord();

  if (!record) {
    showToast("Selecciona una version para imprimir.", "info");
    return;
  }

  document.body.classList.add("print-record-mode");
  window.setTimeout(() => {
    window.print();
  }, 50);
}

function buildCodigoSuggestionLabel(item) {
  const details = [];

  if (item.operaciones) {
    details.push(item.operaciones);
  }

  if (item.tipoMaq) {
    details.push(`Maq: ${item.tipoMaq}`);
  }

  if (item.hasTiempo) {
    details.push(`Tiempo: ${AppUtils.formatNumber(item.tiempoEstimado, 2)}`);
  }

  return details.join(" · ");
}

function hydrateCatalog(payload) {
  appState.catalogByCode.clear();
  appState.puntadasByTipoMaq.clear();

  (payload.basedatos || []).forEach((item) => {
    const normalizedCode = AppUtils.normalizeKey(item.codigo);
    const rawTiempo = item.tiempo ?? item.tiempoEstimado;

    if (!normalizedCode) {
      return;
    }

    appState.catalogByCode.set(normalizedCode, {
      codigo: String(item.codigo ?? "").trim(),
      bloque: String(item.bloque ?? "").trim(),
      operaciones: String(item.operaciones ?? item.descripcionAbreviada ?? "").trim(),
      tiempoEstimado: AppUtils.safeNumber(rawTiempo),
      hasTiempo: rawTiempo !== "" && rawTiempo !== null && rawTiempo !== undefined,
      tipoMaq: String(item.tipoMaq ?? "").trim(),
    });
  });

  (payload.puntadas || []).forEach((item) => {
    const tipoMaq = AppUtils.normalizeKey(item.tipoMaq);

    if (!tipoMaq) {
      return;
    }

    appState.puntadasByTipoMaq.set(tipoMaq, String(item.puntada ?? "").trim());
  });

  dom.codigoSuggestions.innerHTML = "";
  Array.from(appState.catalogByCode.values())
    .sort((left, right) => {
      const leftCode = AppUtils.safeNumber(left.codigo);
      const rightCode = AppUtils.safeNumber(right.codigo);
      const leftIsNumeric = /^\d+(?:[.,]\d+)?$/.test(String(left.codigo ?? "").trim());
      const rightIsNumeric = /^\d+(?:[.,]\d+)?$/.test(String(right.codigo ?? "").trim());

      if (leftIsNumeric && rightIsNumeric && leftCode !== rightCode) {
        return leftCode - rightCode;
      }

      return String(left.codigo ?? "").localeCompare(String(right.codigo ?? ""), "es", {
        numeric: true,
        sensitivity: "base",
      });
    })
    .forEach((item) => {
    const option = document.createElement("option");
    const suggestionLabel = buildCodigoSuggestionLabel(item);
    option.value = item.codigo;
    option.label = `${item.codigo} Â· ${item.operaciones}`;
    dom.codigoSuggestions.appendChild(option);
    option.label = suggestionLabel;
    option.textContent = suggestionLabel;
    option.title = suggestionLabel;
  });
}

function setConnectionState(status, message) {
  if (!dom.connectionBadge) {
    return;
  }

  dom.connectionBadge.className = `status-badge ${status}`;
  dom.connectionBadge.textContent = message;
}

function showLoadingModal(message = "Cargando datos") {
  if (!dom.loadingModal) {
    return;
  }

  document.body.classList.add("is-loading");
  dom.appShell?.setAttribute("inert", "");
  dom.loadingModalTitle.textContent = message;
  dom.loadingModal.classList.remove("hidden");
  dom.loadingModal.setAttribute("aria-hidden", "false");
  dom.loadingModalCard?.focus();
}

function hideLoadingModal() {
  if (!dom.loadingModal) {
    return;
  }

  document.body.classList.remove("is-loading");
  dom.appShell?.removeAttribute("inert");
  dom.loadingModal.classList.add("hidden");
  dom.loadingModal.setAttribute("aria-hidden", "true");
}

function showToast(message, type = "info") {
  window.clearTimeout(toastTimerId);
  dom.toast.hidden = false;
  dom.toast.className = `toast ${type === "success" ? "is-success" : ""} ${
    type === "error" ? "is-error" : ""
  }`.trim();
  dom.toast.textContent = message;

  toastTimerId = window.setTimeout(() => {
    dom.toast.hidden = true;
  }, 3200);
}



