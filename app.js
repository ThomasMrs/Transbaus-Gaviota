import {
  clamp,
  createId,
  escapeAttribute,
  escapeHtml,
  formatDate,
  normalizeFreeText,
  normalizeStoredDate,
  pluralize,
} from "./src/shared.mjs";
import { parseLabelText } from "./src/label-parser.mjs";
import {
  findExistingParcel,
  formatKnownWeightSummary,
  formatRouteCodeForDisplay,
  getDestinationShortLabel,
  getParcelHandlingFactor,
  getParcelIdentifier,
  getParcelCommandNumber,
  getSortingEffortLabel,
  getSortingHandlingAdvice,
  normalizeBarcode,
  normalizeCommandNumber,
  normalizeDestination,
  normalizeDestinationRuleText,
  normalizeParcelData,
  normalizeRouteCode,
  parseDestinationRulePatterns,
  parseParcelWeightKg,
} from "./src/parcel-utils.mjs";
import {
  buildDeliveryEntryDestinationLabel,
  compareDeliveryNoteEntries,
  countIncomparableParcels,
  formatFileSize,
  getDeliveryNoteErrorMessage,
  looksLikePdf,
  normalizeDeliveryNote,
  normalizeDeliveryTextLine,
  parseDeliveryNoteText,
} from "./src/delivery-notes.mjs";
import { createSharedStateStore } from "./src/supabase-shared-state.mjs";

const LEGACY_STATE_STORAGE_KEY = "transbaus-gaviota-state-v1";
const COLLAPSE_STORAGE_KEY = "le-baus-du-tri-collapse-v2";
const ACCESS_RATE_LIMIT_STORAGE_KEY = "transbaus-gaviota-access-rate-v1";
const LEGACY_SHARED_SYNC_META_STORAGE_KEY = "transbaus-gaviota-shared-sync-v1";
const ACCESS_FAILED_ATTEMPTS_LIMIT = 3;
const ACCESS_LOCK_DURATION_MS = 10_000;
const SHARED_SYNC_POLL_MS = 8_000;
const LABEL_AUTO_CAPTURE_POLL_MS = 220;
const LABEL_AUTO_CAPTURE_STABLE_FRAMES = 7;
const LABEL_AUTO_CAPTURE_KICKOFF_MS = 650;
const LABEL_BURST_COUNT = 5;
const LABEL_BURST_INTERVAL_MS = 320;
const PDF_DB_NAME = "le-baus-du-tri-documents-v1";
const PDF_STORE_NAME = "delivery-notes";
const PDFJS_SCRIPT_URL = "vendor/pdf.min.js";
const PDFJS_WORKER_URL = "vendor/pdf.worker.min.js";
const DEFAULT_COLLAPSE_STATE = {
  flow: false,
  scanner: false,
  baqueForm: false,
  smallParcels: false,
  search: false,
  savedPages: false,
  deliveryNote: false,
  destinations: false,
  baques: false,
};
const DEFAULT_BAQUES = [
  { name: "Baque 1", location: "Zone A" },
  { name: "Baque 2", location: "Zone B" },
  { name: "Baque 3", location: "Zone C" },
  { name: "Baque 4", location: "Zone D" },
];

const workspacePage = getWorkspacePageContext();
const state = createDefaultState();
const workspaceLibrary = {
  pages: [],
};
const workspaceEditor = {
  mode: "create",
  targetPageId: "",
};
const workspaceDeleteDialog = {
  targetPageId: "",
};
const collapseState = loadCollapseState();
const accessRateLimit = loadAccessRateLimit();
const sharedSyncMeta = {
  updatedAt: "",
};
const ui = {};
const sharedSync = {
  online: false,
  pollingId: 0,
  requestInFlight: false,
  pendingPush: false,
  offlineToastShown: false,
  initialized: false,
};
const scanner = {
  instance: null,
  active: false,
  handled: false,
  importingBarcode: false,
  target: "parcel",
};
const ocr = {
  worker: null,
  busy: false,
};
const captureSession = {
  stream: null,
  mode: "label",
  busy: false,
  autoCaptureTimer: 0,
  autoCaptureKickoffTimer: 0,
  lastFrameSignature: null,
  stableFrameCount: 0,
  autoTriggered: false,
  analysisCanvas: null,
  analysisContext: null,
};
const deliveryNoteAnalysis = {
  busy: false,
};
let pdfjsLibPromise = null;
let loginLockCountdownId = null;
let sharedStateStore = null;

document.addEventListener("DOMContentLoaded", () => {
  void initializeApp();
});

async function initializeApp() {
  cacheElements();
  bindEvents();
  syncWorkspacePageUi();
  clearLegacyLocalState();
  syncSharedStateBadge("required");
  try {
    sharedStateStore = createSharedStateStore({
      pageId: workspacePage.id,
      pageTitle: () => workspacePage.title,
    });
  } catch (error) {
    console.error("Client Supabase indisponible", error);
    markSharedSyncOffline(error);
  }
  await syncAccessGate();
  render();
}

function getWorkspacePageContext() {
  const params = new URLSearchParams(globalThis.location?.search || "");
  const rawPageId = normalizeWorkspacePageId(params.get("page") || "");
  const requestedTitle = normalizeFreeText(params.get("label") || "");
  if (!rawPageId || rawPageId === "global") {
    return {
      id: "global",
      label: "principale",
      title: "Page principale",
      requestedTitle: "",
      isPrimary: true,
    };
  }

  return {
    id: rawPageId,
    label: requestedTitle || rawPageId.replace(/[-_]+/g, " "),
    title: requestedTitle || rawPageId.replace(/[-_]+/g, " "),
    requestedTitle,
    isPrimary: false,
  };
}

function normalizeWorkspacePageId(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-_]/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^-+|-+$/g, "");
}

function syncWorkspacePageUi() {
  const titleSuffix = workspacePage.isPrimary ? "" : ` - ${workspacePage.title}`;
  document.title = `Le Baus du Tri${titleSuffix}`;

  if (ui.workspaceBadge) {
    ui.workspaceBadge.textContent = workspacePage.isPrimary
      ? "Page principale"
      : `Page ${workspacePage.title}`;
  }
}

function handleNewWorkspaceClick() {
  openWorkspaceCreateModal();
}

function generateWorkspacePageId(seed = "") {
  const now = new Date();
  const datePart = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
  ].join("");
  const timePart = [
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join("");
  const randomPart = Math.random().toString(36).slice(2, 6);
  const slugSeed = normalizeWorkspacePageId(seed).slice(0, 18);
  return `page-${slugSeed || datePart}-${timePart}-${randomPart}`;
}

function buildWorkspacePageUrl(pageId, title = "") {
  const url = new URL(globalThis.location?.href || "http://localhost/");
  const normalizedPageId = normalizeWorkspacePageId(pageId);
  url.pathname = url.pathname.endsWith("/") ? `${url.pathname}index.html` : url.pathname;
  if (!normalizedPageId || normalizedPageId === "global") {
    url.searchParams.delete("page");
    url.searchParams.delete("label");
  } else {
    url.searchParams.set("page", normalizedPageId);
    if (title) {
      url.searchParams.set("label", title);
    } else {
      url.searchParams.delete("label");
    }
  }
  url.hash = "";
  return url.toString();
}

function openWorkspaceCreateModal() {
  openWorkspaceEditorModal({
    mode: "create",
    pageId: "",
    kicker: "Nouvelle page",
    title: "Creer une page vierge",
    help: "La nouvelle page s'ouvrira vide dans un nouvel onglet et restera disponible dans les pages enregistrees.",
    submitLabel: "Creer la page",
    value: getDefaultWorkspaceTitle(),
  });
}

function openWorkspaceRenameModal(pageId) {
  const page = getWorkspaceSummary(pageId);
  if (!page || page.id === "global") {
    showToast("Cette page ne peut pas etre renommee.", "danger");
    return;
  }

  openWorkspaceEditorModal({
    mode: "rename",
    pageId: page.id,
    kicker: "Gestion",
    title: "Renommer la page",
    help: "Le nouveau nom sera visible dans l'historique et sur tous les appareils.",
    submitLabel: "Enregistrer le nom",
    value: page.title || "",
  });
}

function openWorkspaceEditorModal(config) {
  if (!ui.workspaceCreateModal || !ui.workspaceCreateInput) {
    return;
  }

  workspaceEditor.mode = config.mode || "create";
  workspaceEditor.targetPageId = config.pageId || "";
  ui.workspaceCreateForm?.reset();
  if (ui.workspaceCreateKicker) {
    ui.workspaceCreateKicker.textContent = config.kicker || "Nouvelle page";
  }
  if (ui.workspaceCreateTitle) {
    ui.workspaceCreateTitle.textContent = config.title || "Creer une page vierge";
  }
  if (ui.workspaceCreateHelp) {
    ui.workspaceCreateHelp.textContent = config.help || "";
  }
  if (ui.workspaceCreateSubmitLabel) {
    ui.workspaceCreateSubmitLabel.textContent = config.submitLabel || "Valider";
  }
  ui.workspaceCreateInput.value = config.value || "";
  ui.workspaceCreateModal.classList.remove("hidden");
  ui.workspaceCreateModal.setAttribute("aria-hidden", "false");
  window.setTimeout(() => {
    ui.workspaceCreateInput.focus();
    ui.workspaceCreateInput.select();
  }, 0);
}

function closeWorkspaceCreateModal() {
  if (!ui.workspaceCreateModal) {
    return;
  }

  workspaceEditor.mode = "create";
  workspaceEditor.targetPageId = "";
  ui.workspaceCreateModal.classList.add("hidden");
  ui.workspaceCreateModal.setAttribute("aria-hidden", "true");
}

function getDefaultWorkspaceTitle() {
  return `Page ${new Date().toLocaleDateString("fr-FR")}`;
}

async function handleWorkspaceCreateSubmit(event) {
  event.preventDefault();

  const normalizedTitle = normalizeFreeText(ui.workspaceCreateInput?.value || "") || getDefaultWorkspaceTitle();
  if (workspaceEditor.mode === "rename") {
    await renameWorkspacePage(workspaceEditor.targetPageId, normalizedTitle);
    return;
  }

  const pageId = generateWorkspacePageId(normalizedTitle);
  closeWorkspaceCreateModal();
  window.open(buildWorkspacePageUrl(pageId, normalizedTitle), "_blank", "noopener,noreferrer");
}

function openWorkspaceDeleteModal(pageId) {
  const page = getWorkspaceSummary(pageId);
  if (!page || page.id === "global" || !ui.workspaceDeleteModal) {
    showToast("Cette page ne peut pas etre supprimee.", "danger");
    return;
  }

  workspaceDeleteDialog.targetPageId = page.id;
  if (ui.workspaceDeleteMessage) {
    ui.workspaceDeleteMessage.textContent = `Supprimer "${page.title || "Page sans nom"}" de l'historique ? Cette action efface ses colis, ses baques et ses regles.`;
  }
  ui.workspaceDeleteModal.classList.remove("hidden");
  ui.workspaceDeleteModal.setAttribute("aria-hidden", "false");
}

function closeWorkspaceDeleteModal() {
  if (!ui.workspaceDeleteModal) {
    return;
  }

  workspaceDeleteDialog.targetPageId = "";
  ui.workspaceDeleteModal.classList.add("hidden");
  ui.workspaceDeleteModal.setAttribute("aria-hidden", "true");
}

function getWorkspaceSummary(pageId) {
  const normalizedPageId = normalizeWorkspacePageId(pageId);
  if (!normalizedPageId && pageId !== "global") {
    return null;
  }

  const resolvedPageId = normalizedPageId || "global";
  const existingPage = workspaceLibrary.pages.find((page) => page.id === resolvedPageId);
  if (existingPage) {
    return existingPage;
  }

  if (resolvedPageId === workspacePage.id) {
    return {
      id: workspacePage.id,
      title: workspacePage.title,
      createdAt: "",
      updatedAt: "",
      parcelsCount: state.parcels.length + getSmallParcelCountTotal(),
      baquesCount: state.baques.length,
    };
  }

  return null;
}

function syncCurrentWorkspaceFromLibrary() {
  if (workspacePage.isPrimary) {
    return;
  }

  const currentPage = getWorkspaceSummary(workspacePage.id);
  if (!currentPage) {
    return;
  }

  const nextTitle = normalizeFreeText(currentPage.title || "");
  if (!nextTitle || nextTitle === workspacePage.title) {
    return;
  }

  workspacePage.title = nextTitle;
  workspacePage.label = nextTitle;
  workspacePage.requestedTitle = nextTitle;
  syncWorkspacePageUi();
  window.history.replaceState({}, "", buildWorkspacePageUrl(workspacePage.id, nextTitle));
}

async function waitForSharedSyncIdle(timeoutMs = 5000) {
  const deadline = Date.now() + timeoutMs;
  while (sharedSync.requestInFlight && Date.now() < deadline) {
    await new Promise((resolve) => {
      window.setTimeout(resolve, 120);
    });
  }

  if (sharedSync.requestInFlight) {
    const error = new Error("shared-sync-busy");
    error.code = "shared-sync-busy";
    throw error;
  }
}

function getWorkspaceActionErrorMessage(error, fallbackMessage) {
  const code = String(error?.code || "");
  if (code === "workspace-primary-protected") {
    return "La page principale ne peut pas etre modifiee ici.";
  }

  if (code === "workspace-page-missing") {
    return "Cette page n'existe plus dans l'historique. Rechargez la liste.";
  }

  if (code === "shared-sync-busy") {
    return "Une synchronisation est deja en cours. Reessayez dans quelques secondes.";
  }

  return fallbackMessage;
}

async function renameWorkspacePage(pageId, nextTitle) {
  if (!sharedStateStore) {
    showToast("La BDD cloud est indisponible pour l'instant.", "danger");
    return;
  }

  const page = getWorkspaceSummary(pageId);
  if (!page || page.id === "global") {
    showToast("Cette page ne peut pas etre renommee.", "danger");
    return;
  }

  const normalizedTitle = normalizeFreeText(nextTitle || "") || page.title || getDefaultWorkspaceTitle();
  if (normalizedTitle === page.title) {
    closeWorkspaceCreateModal();
    return;
  }

  try {
    await waitForSharedSyncIdle();
    if (sharedSync.pendingPush) {
      await flushSharedStateSyncQueue();
    }

    sharedSync.requestInFlight = true;
    const payload = await sharedStateStore.renamePage(page.id, normalizedTitle);
    sharedSyncMeta.updatedAt = normalizeStoredDate(payload.updatedAt || "", new Date().toISOString());
    workspaceLibrary.pages = payload.pages || workspaceLibrary.pages;
    if (page.id === workspacePage.id) {
      workspacePage.title = normalizedTitle;
      workspacePage.label = normalizedTitle;
      workspacePage.requestedTitle = normalizedTitle;
      syncWorkspacePageUi();
      window.history.replaceState({}, "", buildWorkspacePageUrl(workspacePage.id, normalizedTitle));
    } else {
      syncCurrentWorkspaceFromLibrary();
    }
    closeWorkspaceCreateModal();
    renderWorkspacePages();
    markSharedSyncOnline();
    showToast(`Page "${normalizedTitle}" renommee.`);
  } catch (error) {
    console.error("Renommage de page impossible", error);
    if (/^(workspace-|shared-sync-busy)/.test(String(error?.code || ""))) {
      showToast(getWorkspaceActionErrorMessage(error, "Impossible de renommer cette page pour le moment."), "danger");
    } else {
      markSharedSyncOffline(error);
    }
  } finally {
    sharedSync.requestInFlight = false;
  }
}

async function deleteWorkspacePage(pageId) {
  if (!sharedStateStore) {
    showToast("La BDD cloud est indisponible pour l'instant.", "danger");
    return;
  }

  const page = getWorkspaceSummary(pageId);
  if (!page || page.id === "global") {
    showToast("Cette page ne peut pas etre supprimee.", "danger");
    return;
  }

  try {
    await waitForSharedSyncIdle();
    if (sharedSync.pendingPush) {
      await flushSharedStateSyncQueue();
    }

    sharedSync.requestInFlight = true;
    const payload = await sharedStateStore.deletePage(page.id);
    sharedSyncMeta.updatedAt = normalizeStoredDate(payload.updatedAt || "", new Date().toISOString());
    workspaceLibrary.pages = payload.pages || workspaceLibrary.pages;
    closeWorkspaceDeleteModal();
    markSharedSyncOnline();
    if (page.id === workspacePage.id) {
      window.location.assign(buildWorkspacePageUrl("", ""));
      return;
    }

    renderWorkspacePages();
    showToast(`Page "${page.title}" supprimee.`);
  } catch (error) {
    console.error("Suppression de page impossible", error);
    if (/^(workspace-|shared-sync-busy)/.test(String(error?.code || ""))) {
      showToast(getWorkspaceActionErrorMessage(error, "Impossible de supprimer cette page pour le moment."), "danger");
    } else {
      markSharedSyncOffline(error);
    }
  } finally {
    sharedSync.requestInFlight = false;
  }
}

async function handleWorkspaceDeleteConfirm() {
  await deleteWorkspacePage(workspaceDeleteDialog.targetPageId);
}

async function handleWorkspaceListClick(event) {
  const actionButton = event.target.closest("[data-workspace-action]");
  if (!(actionButton instanceof HTMLButtonElement)) {
    return;
  }

  const pageId = actionButton.dataset.pageId || "";
  const action = actionButton.dataset.workspaceAction || "";
  if (!pageId || !action) {
    return;
  }

  if (action === "rename-workspace") {
    openWorkspaceRenameModal(pageId);
    return;
  }

  if (action === "delete-workspace") {
    openWorkspaceDeleteModal(pageId);
  }
}

function cacheElements() {
  ui.loginGate = document.querySelector("#loginGate");
  ui.loginForm = document.querySelector("#loginForm");
  ui.loginPasswordInput = document.querySelector("#loginPasswordInput");
  ui.toggleLoginPasswordBtn = document.querySelector("#toggleLoginPasswordBtn");
  ui.toggleLoginPasswordIcon = document.querySelector("#toggleLoginPasswordIcon");
  ui.loginStatus = document.querySelector("#loginStatus");
  ui.loginSubmitBtn = ui.loginForm?.querySelector('button[type="submit"]');
  ui.logoutBtn = document.querySelector("#logoutBtn");
  ui.newWorkspaceBtn = document.querySelector("#newWorkspaceBtn");
  ui.workspaceBadge = document.querySelector("#workspaceBadge");
  ui.workspaceList = document.querySelector("#workspaceList");
  ui.workspaceCreateModal = document.querySelector("#workspaceCreateModal");
  ui.workspaceCreateForm = document.querySelector("#workspaceCreateForm");
  ui.workspaceCreateKicker = document.querySelector("#workspaceCreateKicker");
  ui.workspaceCreateTitle = document.querySelector("#workspaceCreateTitle");
  ui.workspaceCreateInput = document.querySelector("#workspaceCreateInput");
  ui.workspaceCreateHelp = document.querySelector("#workspaceCreateHelp");
  ui.workspaceCreateSubmitLabel = document.querySelector("#workspaceCreateSubmitLabel");
  ui.closeWorkspaceCreateBtn = document.querySelector("#closeWorkspaceCreateBtn");
  ui.cancelWorkspaceCreateBtn = document.querySelector("#cancelWorkspaceCreateBtn");
  ui.workspaceDeleteModal = document.querySelector("#workspaceDeleteModal");
  ui.workspaceDeleteMessage = document.querySelector("#workspaceDeleteMessage");
  ui.closeWorkspaceDeleteBtn = document.querySelector("#closeWorkspaceDeleteBtn");
  ui.cancelWorkspaceDeleteBtn = document.querySelector("#cancelWorkspaceDeleteBtn");
  ui.confirmWorkspaceDeleteBtn = document.querySelector("#confirmWorkspaceDeleteBtn");
  ui.syncStatusBadge = document.querySelector("#syncStatusBadge");
  ui.heroStats = document.querySelector("#heroStats");
  ui.parcelForm = document.querySelector("#parcelForm");
  ui.parcelBaqueSelect = document.querySelector("#parcelBaqueSelect");
  ui.routeCodeInput = document.querySelector("#routeCodeInput");
  ui.destinationInput = document.querySelector("#destinationInput");
  ui.clientInput = document.querySelector("#clientInput");
  ui.descriptionInput = document.querySelector("#descriptionInput");
  ui.routeLabelInput = document.querySelector("#routeLabelInput");
  ui.referenceInput = document.querySelector("#referenceInput");
  ui.shippingDateInput = document.querySelector("#shippingDateInput");
  ui.weightInput = document.querySelector("#weightInput");
  ui.measuredDimensionsInput = document.querySelector("#measuredDimensionsInput");
  ui.packageIndexInput = document.querySelector("#packageIndexInput");
  ui.barcodeInput = document.querySelector("#barcodeInput");
  ui.openScannerBtn = document.querySelector("#openScannerBtn");
  ui.openSmallParcelScannerBtn = document.querySelector("#openSmallParcelScannerBtn");
  ui.importBarcodeBtn = document.querySelector("#importBarcodeBtn");
  ui.chooseBarcodeBtn = document.querySelector("#chooseBarcodeBtn");
  ui.scanLabelBtn = document.querySelector("#scanLabelBtn");
  ui.chooseLabelBtn = document.querySelector("#chooseLabelBtn");
  ui.labelCameraInput = document.querySelector("#labelCameraInput");
  ui.labelLibraryInput = document.querySelector("#labelLibraryInput");
  ui.barcodeCameraInput = document.querySelector("#barcodeCameraInput");
  ui.barcodeLibraryInput = document.querySelector("#barcodeLibraryInput");
  ui.ocrStatus = document.querySelector("#ocrStatus");
  ui.barcodeStatus = document.querySelector("#barcodeStatus");
  ui.baqueForm = document.querySelector("#baqueForm");
  ui.baqueNameInput = document.querySelector("#baqueNameInput");
  ui.baqueLocationInput = document.querySelector("#baqueLocationInput");
  ui.smallParcelForm = document.querySelector("#smallParcelForm");
  ui.smallParcelBarcodeInput = document.querySelector("#smallParcelBarcodeInput");
  ui.smallParcelQuantityInput = document.querySelector("#smallParcelQuantityInput");
  ui.smallParcelSummary = document.querySelector("#smallParcelSummary");
  ui.smallParcelList = document.querySelector("#smallParcelList");
  ui.searchInput = document.querySelector("#searchInput");
  ui.searchResults = document.querySelector("#searchResults");
  ui.importDeliveryNoteBtn = document.querySelector("#importDeliveryNoteBtn");
  ui.deliveryNoteInput = document.querySelector("#deliveryNoteInput");
  ui.deliveryNoteStatus = document.querySelector("#deliveryNoteStatus");
  ui.deliveryNoteList = document.querySelector("#deliveryNoteList");
  ui.destinationRuleForm = document.querySelector("#destinationRuleForm");
  ui.destinationRuleLabelInput = document.querySelector("#destinationRuleLabelInput");
  ui.destinationRuleMatchModeSelect = document.querySelector("#destinationRuleMatchModeSelect");
  ui.destinationRuleTargetBaqueSelect = document.querySelector("#destinationRuleTargetBaqueSelect");
  ui.destinationRulePatternsInput = document.querySelector("#destinationRulePatternsInput");
  ui.destinationRulesList = document.querySelector("#destinationRulesList");
  ui.destinationSummary = document.querySelector("#destinationSummary");
  ui.sortingPlan = document.querySelector("#sortingPlan");
  ui.baquesGrid = document.querySelector("#baquesGrid");
  ui.scannerModal = document.querySelector("#scannerModal");
  ui.reader = document.querySelector("#reader");
  ui.scannerHelp = document.querySelector(".scanner-help");
  ui.scannerStatus = document.querySelector("#scannerStatus");
  ui.closeScannerBtn = document.querySelector("#closeScannerBtn");
  ui.captureModal = document.querySelector("#captureModal");
  ui.capturePreview = document.querySelector(".capture-preview");
  ui.captureVideo = document.querySelector("#captureVideo");
  ui.captureGuide = document.querySelector("#captureGuide");
  ui.captureTitle = document.querySelector("#captureTitle");
  ui.captureHint = document.querySelector("#captureHint");
  ui.captureStatus = document.querySelector("#captureStatus");
  ui.takeCaptureBtn = document.querySelector("#takeCaptureBtn");
  ui.closeCaptureBtn = document.querySelector("#closeCaptureBtn");
  ui.toastZone = document.querySelector("#toastZone");
  ui.barcodeFileReader = document.querySelector("#barcodeFileReader");
}

function bindEvents() {
  document.addEventListener("click", handleCollapseToggle);
  ui.loginForm.addEventListener("submit", handleLoginSubmit);
  ui.toggleLoginPasswordBtn?.addEventListener("click", toggleLoginPasswordVisibility);
  ui.logoutBtn.addEventListener("click", handleLogoutClick);
  ui.newWorkspaceBtn?.addEventListener("click", handleNewWorkspaceClick);
  ui.workspaceCreateForm?.addEventListener("submit", handleWorkspaceCreateSubmit);
  ui.closeWorkspaceCreateBtn?.addEventListener("click", closeWorkspaceCreateModal);
  ui.cancelWorkspaceCreateBtn?.addEventListener("click", closeWorkspaceCreateModal);
  ui.workspaceList?.addEventListener("click", (event) => {
    void handleWorkspaceListClick(event);
  });
  ui.closeWorkspaceDeleteBtn?.addEventListener("click", closeWorkspaceDeleteModal);
  ui.cancelWorkspaceDeleteBtn?.addEventListener("click", closeWorkspaceDeleteModal);
  ui.confirmWorkspaceDeleteBtn?.addEventListener("click", () => {
    void handleWorkspaceDeleteConfirm();
  });
  ui.parcelForm.addEventListener("submit", handleParcelSubmit);
  ui.baqueForm.addEventListener("submit", handleBaqueSubmit);
  ui.smallParcelForm?.addEventListener("submit", handleSmallParcelSubmit);
  ui.smallParcelList?.addEventListener("click", handleSmallParcelListClick);
  ui.searchInput.addEventListener("input", renderSearchResults);
  ui.importDeliveryNoteBtn.addEventListener("click", openDeliveryNotePicker);
  ui.deliveryNoteInput.addEventListener("change", (event) => {
    void handleDeliveryNoteImport(event);
  });
  ui.deliveryNoteList.addEventListener("click", (event) => {
    void handleDeliveryNoteListClick(event);
  });
  ui.destinationRuleForm.addEventListener("submit", handleDestinationRuleSubmit);
  ui.destinationRulesList.addEventListener("click", handleDestinationRulesClick);
  ui.destinationRulesList.addEventListener("change", handleDestinationRulesChange);
  ui.destinationSummary.addEventListener("click", handleDestinationSummaryClick);
  ui.openScannerBtn?.addEventListener("click", openScanner);
  ui.openSmallParcelScannerBtn?.addEventListener("click", () => {
    void openScanner("small-parcel");
  });
  ui.importBarcodeBtn?.addEventListener("click", openBarcodeCameraPicker);
  ui.chooseBarcodeBtn?.addEventListener("click", openBarcodeLibraryPicker);
  ui.scanLabelBtn.addEventListener("click", openLabelCameraPicker);
  ui.chooseLabelBtn.addEventListener("click", openLabelLibraryPicker);
  ui.barcodeCameraInput.addEventListener("change", handleBarcodeImageChange);
  ui.barcodeLibraryInput.addEventListener("change", handleBarcodeImageChange);
  ui.labelCameraInput.addEventListener("change", handleLabelImageChange);
  ui.labelLibraryInput.addEventListener("change", handleLabelImageChange);
  ui.closeScannerBtn.addEventListener("click", closeScanner);
  ui.closeCaptureBtn.addEventListener("click", closeCaptureModal);
  ui.takeCaptureBtn.addEventListener("click", handleCapturePhoto);
  ui.scannerModal.addEventListener("click", handleModalClick);
  ui.captureModal.addEventListener("click", handleModalClick);
  ui.workspaceCreateModal?.addEventListener("click", handleModalClick);
  ui.workspaceDeleteModal?.addEventListener("click", handleModalClick);
  ui.baquesGrid.addEventListener("click", handleBaqueGridClick);
  ui.baquesGrid.addEventListener("change", handleBaqueGridChange);
  window.addEventListener("beforeunload", (event) => {
    if (sharedSync.pendingPush) {
      event.preventDefault();
      event.returnValue = "";
    }

    void stopScanner();
    void stopCaptureStream();
    void stopOcrWorker();
  });
  window.addEventListener("resize", applyCollapseStateToDom);
}

async function syncAccessGate() {
  refreshAccessRateLimit();
  if (!sharedStateStore?.getAccessSession) {
    setAppAccess(false);
    return false;
  }

  try {
    const session = await sharedStateStore.getAccessSession();
    const isGranted = Boolean(session);
    setAppAccess(isGranted);
    if (isGranted) {
      await initializeSharedStateSync();
    }
    return isGranted;
  } catch (error) {
    console.error("Session Supabase indisponible", error);
    setAppAccess(false);
    markSharedSyncOffline(error);
    return false;
  }
}

async function handleLoginSubmit(event) {
  event.preventDefault();

  try {
    refreshAccessRateLimit();
    if (isAccessTemporarilyLocked()) {
      syncAccessRateLimitUi();
      return;
    }

    const typedPassword = ui.loginPasswordInput.value.trim();
    if (!typedPassword) {
      ui.loginStatus.textContent = "Saisissez le code d'acces.";
      ui.loginPasswordInput.focus();
      return;
    }

    if (!sharedStateStore?.signInWithPassword) {
      ui.loginStatus.textContent = "Connexion a Supabase indisponible. Rechargez la page puis reessayez.";
      showToast("Connexion a Supabase indisponible. Rechargez la page puis reessayez.", "danger");
      return;
    }

    await sharedStateStore.signInWithPassword(typedPassword);
    resetAccessRateLimit();
    ui.loginStatus.textContent = "";
    ui.loginForm.reset();
    setAppAccess(true);
    await initializeSharedStateSync({ force: true });
    render();
  } catch (error) {
    console.error("Connexion impossible", error);
    if (/invalid login credentials|invalid_credentials|email not confirmed/i.test(String(error?.message || ""))) {
      registerFailedLoginAttempt();
      if (!isAccessTemporarilyLocked()) {
        ui.loginPasswordInput.focus();
        ui.loginPasswordInput.select();
      }
      return;
    }

    const errorMessage = getSharedSyncErrorMessage(error);
    ui.loginStatus.textContent = errorMessage;
    showToast(errorMessage, "danger");
  }
}

async function handleLogoutClick() {
  try {
    await sharedStateStore?.signOut?.();
  } catch (error) {
    console.error("Deconnexion impossible", error);
  }

  void closeScanner();
  void closeCaptureModal();
  stopSharedStatePolling();
  sharedSync.initialized = false;
  sharedSyncMeta.updatedAt = "";
  workspaceLibrary.pages = [];
  hydrateState(createDefaultState());
  render();
  setAppAccess(false);
}

function setAppAccess(isGranted) {
  document.body.classList.toggle("app-locked", !isGranted);
  ui.loginGate.setAttribute("aria-hidden", String(isGranted));
  ui.logoutBtn.hidden = !isGranted;

  if (!isGranted) {
    syncSharedStateBadge("required");
    syncLoginPasswordVisibility(false);
    ui.loginForm.reset();
    syncAccessRateLimitUi();
    if (!isAccessTemporarilyLocked()) {
      ui.loginPasswordInput.focus();
    }
    return;
  }

  stopAccessLockCountdown();
  syncSharedStateBadge(sharedSync.online ? "shared" : "connecting");
  ui.routeCodeInput.focus();
}

function loadAccessRateLimit() {
  try {
    const raw = window.localStorage.getItem(ACCESS_RATE_LIMIT_STORAGE_KEY);
    if (!raw) {
      return {
        failedAttempts: 0,
        lockedUntil: 0,
      };
    }

    const parsed = JSON.parse(raw);
    return {
      failedAttempts: Math.max(0, Number(parsed.failedAttempts || 0)),
      lockedUntil: Math.max(0, Number(parsed.lockedUntil || 0)),
    };
  } catch (error) {
    return {
      failedAttempts: 0,
      lockedUntil: 0,
    };
  }
}

function saveAccessRateLimit() {
  window.localStorage.setItem(ACCESS_RATE_LIMIT_STORAGE_KEY, JSON.stringify(accessRateLimit));
}

function refreshAccessRateLimit() {
  if (accessRateLimit.lockedUntil && accessRateLimit.lockedUntil <= Date.now()) {
    accessRateLimit.lockedUntil = 0;
    saveAccessRateLimit();
  }
}

function resetAccessRateLimit() {
  accessRateLimit.failedAttempts = 0;
  accessRateLimit.lockedUntil = 0;
  saveAccessRateLimit();
  stopAccessLockCountdown();
  syncAccessRateLimitUi();
}

function isAccessTemporarilyLocked() {
  return accessRateLimit.lockedUntil > Date.now();
}

function registerFailedLoginAttempt() {
  accessRateLimit.failedAttempts += 1;

  if (accessRateLimit.failedAttempts % ACCESS_FAILED_ATTEMPTS_LIMIT === 0) {
    accessRateLimit.lockedUntil = Date.now() + ACCESS_LOCK_DURATION_MS;
    saveAccessRateLimit();
    syncAccessRateLimitUi();
    return;
  }

  const remainingAttempts = ACCESS_FAILED_ATTEMPTS_LIMIT - (accessRateLimit.failedAttempts % ACCESS_FAILED_ATTEMPTS_LIMIT);
  accessRateLimit.lockedUntil = 0;
  saveAccessRateLimit();
  stopAccessLockCountdown();
  ui.loginStatus.textContent = `Code incorrect. Encore ${remainingAttempts} essai${remainingAttempts > 1 ? "s" : ""} avant l'attente.`;
  ui.loginPasswordInput.disabled = false;
  ui.loginSubmitBtn.disabled = false;
}

function syncAccessRateLimitUi() {
  refreshAccessRateLimit();

  const isLocked = isAccessTemporarilyLocked();
  ui.loginPasswordInput.disabled = isLocked;
  ui.loginSubmitBtn.disabled = isLocked;

  if (isLocked) {
    ui.loginStatus.textContent = formatAccessLockMessage();
    startAccessLockCountdown();
    return;
  }

  stopAccessLockCountdown();
  if (isAccessLockMessage(ui.loginStatus.textContent)) {
    ui.loginStatus.textContent = "";
  }
}

function startAccessLockCountdown() {
  if (loginLockCountdownId !== null) {
    return;
  }

  loginLockCountdownId = window.setInterval(() => {
    refreshAccessRateLimit();

    if (!isAccessTemporarilyLocked()) {
      syncAccessRateLimitUi();
      if (ui.loginGate.getAttribute("aria-hidden") !== "true") {
        ui.loginPasswordInput.focus();
      }
      return;
    }

    ui.loginStatus.textContent = formatAccessLockMessage();
  }, 250);
}

function stopAccessLockCountdown() {
  if (loginLockCountdownId === null) {
    return;
  }

  window.clearInterval(loginLockCountdownId);
  loginLockCountdownId = null;
}

function formatAccessLockMessage() {
  const remainingSeconds = Math.max(1, Math.ceil((accessRateLimit.lockedUntil - Date.now()) / 1000));
  return `Trop d'essais rates. Reessayez dans ${remainingSeconds} seconde${remainingSeconds > 1 ? "s" : ""}.`;
}

function isAccessLockMessage(value) {
  return /^Trop d'essais rates\./.test(value);
}

function toggleLoginPasswordVisibility() {
  const shouldShow = ui.loginPasswordInput?.type === "password";
  syncLoginPasswordVisibility(shouldShow);
}

function syncLoginPasswordVisibility(isVisible) {
  if (!ui.loginPasswordInput || !ui.toggleLoginPasswordBtn) {
    return;
  }

  ui.loginPasswordInput.type = isVisible ? "text" : "password";
  ui.toggleLoginPasswordBtn.setAttribute("aria-pressed", String(isVisible));
  ui.toggleLoginPasswordBtn.setAttribute("aria-label", isVisible ? "Masquer le code" : "Afficher le code");
  ui.toggleLoginPasswordBtn.setAttribute("title", isVisible ? "Masquer le code" : "Afficher le code");
  if (ui.toggleLoginPasswordIcon) {
    ui.toggleLoginPasswordIcon.innerHTML = isVisible
      ? `
        <path
          d="M3 4.5 21 19.5"
          fill="none"
          stroke="currentColor"
          stroke-linecap="round"
          stroke-linejoin="round"
          stroke-width="1.8"
        />
        <path
          d="M10.7 6.3A10.72 10.72 0 0 1 12 6c6 0 9.75 6 9.75 6a18.88 18.88 0 0 1-4.03 4.49M6.44 8.2C4.11 9.72 2.25 12 2.25 12s3.75 6 9.75 6c1.51 0 2.89-.38 4.12-.97"
          fill="none"
          stroke="currentColor"
          stroke-linecap="round"
          stroke-linejoin="round"
          stroke-width="1.8"
        />
        <path
          d="M9.88 9.88A3.01 3.01 0 0 0 9 12c0 1.66 1.34 3 3 3 .78 0 1.49-.3 2.02-.8"
          fill="none"
          stroke="currentColor"
          stroke-linecap="round"
          stroke-linejoin="round"
          stroke-width="1.8"
        />
      `
      : `
        <path
          d="M2.25 12s3.75-6 9.75-6 9.75 6 9.75 6-3.75 6-9.75 6-9.75-6-9.75-6Z"
          fill="none"
          stroke="currentColor"
          stroke-linecap="round"
          stroke-linejoin="round"
          stroke-width="1.8"
        />
        <circle
          cx="12"
          cy="12"
          r="3.2"
          fill="none"
          stroke="currentColor"
          stroke-width="1.8"
        />
      `;
  }
}

function clearLegacyLocalState() {
  window.localStorage.removeItem(LEGACY_STATE_STORAGE_KEY);
  window.localStorage.removeItem(LEGACY_SHARED_SYNC_META_STORAGE_KEY);
  window.localStorage.removeItem("transbaus-gaviota-access-v1");
}

function normalizeSmallParcelScan(scan) {
  const normalizedCode = normalizeBarcode(scan?.barcode || scan?.commandNumber || "");
  const normalizedParcelData = normalizeParcelData({
    barcode: normalizedCode,
    commandNumber: normalizeCommandNumber(scan?.commandNumber || normalizedCode),
  });
  const commandNumber = getParcelCommandNumber(normalizedParcelData);
  if (!commandNumber) {
    return null;
  }

  const quantity = clamp(Math.round(Number(scan?.quantity || 1)), 1, 99);
  const createdAt = normalizeStoredDate(scan?.createdAt || "", new Date().toISOString());
  return {
    id: scan?.id || createId(),
    barcode: normalizedParcelData.barcode || commandNumber,
    commandNumber,
    quantity,
    createdAt,
    updatedAt: normalizeStoredDate(scan?.updatedAt || createdAt, createdAt),
  };
}

function normalizePersistedState(parsed) {
  if (!Array.isArray(parsed?.baques) || !Array.isArray(parsed?.parcels)) {
    return createDefaultState();
  }

  const fallbackState = createDefaultState();
  const baques = parsed.baques.map((baque) => ({
    id: baque.id || createId(),
    name: String(baque.name || "Baque"),
    location: String(baque.location || "Sans emplacement"),
    validatedAt: normalizeStoredDate(baque.validatedAt || "", ""),
    createdAt: normalizeStoredDate(baque.createdAt, new Date().toISOString()),
    updatedAt: normalizeStoredDate(baque.updatedAt || baque.validatedAt || baque.createdAt, new Date().toISOString()),
  }));

  const availableBaqueIds = new Set(baques.map((baque) => baque.id));
  const parcels = parsed.parcels
    .filter((parcel) => parcel && availableBaqueIds.has(parcel.currentBaqueId))
    .map((parcel) => normalizeParcelData({
      id: parcel.id || createId(),
      barcode: String(parcel.barcode || "").trim(),
      commandNumber: String(parcel.commandNumber || "").trim(),
      routeCode: String(parcel.routeCode || "").trim().toUpperCase(),
      destination: String(parcel.destination || "").trim(),
      client: String(parcel.client || "").trim(),
      description: String(parcel.description || "").trim(),
      routeLabel: String(parcel.routeLabel || "").trim(),
      reference: String(parcel.reference || "").trim(),
      shippingDate: String(parcel.shippingDate || "").trim(),
      weight: String(parcel.weight || "").trim(),
      measuredDimensions: String(parcel.measuredDimensions || "").trim(),
      packageIndex: String(parcel.packageIndex || "").trim(),
      currentBaqueId: parcel.currentBaqueId,
      originBaqueId: parcel.originBaqueId || parcel.currentBaqueId,
      originBaqueLabel: String(parcel.originBaqueLabel || ""),
      createdAt: normalizeStoredDate(parcel.createdAt, new Date().toISOString()),
      updatedAt: normalizeStoredDate(parcel.updatedAt || parcel.createdAt, new Date().toISOString()),
    }))
    .filter((parcel) => parcel.routeCode || parcel.commandNumber || parcel.barcode || parcel.destination);
  const deliveryNotes = Array.isArray(parsed.deliveryNotes)
    ? parsed.deliveryNotes
      .map((note) => normalizeDeliveryNote(note))
      .filter(Boolean)
    : [];
  const destinationRules = Array.isArray(parsed.destinationRules)
    ? parsed.destinationRules
      .map((rule) => normalizeDestinationRule(rule, { availableBaqueIds }))
      .filter(Boolean)
    : [];
  const smallParcelScans = Array.isArray(parsed.smallParcelScans)
    ? parsed.smallParcelScans
      .map((scan) => normalizeSmallParcelScan(scan))
      .filter(Boolean)
    : [];

  return {
    baques: baques.length ? baques : fallbackState.baques,
    parcels,
    smallParcelScans,
    deliveryNotes,
    destinationRules,
  };
}

function hydrateState(nextState) {
  const normalizedState = normalizePersistedState(nextState);
  state.baques.splice(0, state.baques.length, ...normalizedState.baques);
  state.parcels.splice(0, state.parcels.length, ...normalizedState.parcels);
  state.smallParcelScans.splice(0, state.smallParcelScans.length, ...normalizedState.smallParcelScans);
  state.deliveryNotes.splice(0, state.deliveryNotes.length, ...normalizedState.deliveryNotes);
  state.destinationRules.splice(0, state.destinationRules.length, ...normalizedState.destinationRules);
  return normalizedState;
}

function createDefaultState() {
  const now = new Date().toISOString();
  return {
    baques: DEFAULT_BAQUES.map((baque) => ({
      id: createId(),
      name: baque.name,
      location: baque.location,
      validatedAt: "",
      createdAt: now,
      updatedAt: now,
    })),
    parcels: [],
    smallParcelScans: [],
    deliveryNotes: [],
    destinationRules: [],
  };
}

function loadCollapseState() {
  try {
    const raw = window.localStorage.getItem(COLLAPSE_STORAGE_KEY);
    if (!raw) {
      return { ...DEFAULT_COLLAPSE_STATE };
    }

    const parsed = JSON.parse(raw);
    return Object.fromEntries(
      Object.keys(DEFAULT_COLLAPSE_STATE).map((key) => [key, Boolean(parsed?.[key] ?? DEFAULT_COLLAPSE_STATE[key])]),
    );
  } catch (error) {
    return { ...DEFAULT_COLLAPSE_STATE };
  }
}

function saveCollapseState() {
  window.localStorage.setItem(COLLAPSE_STORAGE_KEY, JSON.stringify(collapseState));
}

function saveState() {
  refreshStoredDeliveryNoteAnalyses();
  queueSharedStateSync();
}

function syncSharedStateBadge(mode) {
  if (!ui.syncStatusBadge) {
    return;
  }

  const labels = {
    connecting: "Connexion BDD...",
    offline: "BDD hors ligne",
    shared: "BDD partagee",
    required: "BDD requise",
  };

  ui.syncStatusBadge.dataset.syncState = mode;
  ui.syncStatusBadge.textContent = labels[mode] || labels.required;
}

async function initializeSharedStateSync(options = {}) {
  if (!sharedStateStore || (sharedSync.initialized && !options.force)) {
    return;
  }

  syncSharedStateBadge("connecting");
  const record = await fetchSharedStateRecord();
  workspaceLibrary.pages = record?.pages || [];
  syncCurrentWorkspaceFromLibrary();
  if (record?.state) {
    hydrateState(record.state);
    if (refreshStoredDeliveryNoteAnalyses()) {
      queueSharedStateSync();
    }
    sharedSyncMeta.updatedAt = record.updatedAt;
  } else if (record) {
    if (!workspacePage.isPrimary) {
      sharedSync.pendingPush = true;
      await flushSharedStateSyncQueue();
    }
  }

  sharedSync.initialized = true;
  startSharedStatePolling();
}

function startSharedStatePolling() {
  if (sharedSync.pollingId) {
    return;
  }

  sharedSync.pollingId = window.setInterval(() => {
    void pollSharedState();
  }, SHARED_SYNC_POLL_MS);
}

function stopSharedStatePolling() {
  if (!sharedSync.pollingId) {
    return;
  }

  window.clearInterval(sharedSync.pollingId);
  sharedSync.pollingId = 0;
}

async function pollSharedState() {
  if (sharedSync.requestInFlight) {
    return;
  }

  if (sharedSync.pendingPush) {
    await flushSharedStateSyncQueue();
    return;
  }

  await pullSharedStateFromServer();
}

function queueSharedStateSync() {
  sharedSync.pendingPush = true;
  if (sharedSync.online) {
    void flushSharedStateSyncQueue();
  }
}

async function flushSharedStateSyncQueue() {
  if (!sharedStateStore || sharedSync.requestInFlight || !sharedSync.pendingPush) {
    return;
  }

  sharedSync.requestInFlight = true;
  sharedSync.pendingPush = false;

  try {
    const payload = await sharedStateStore.saveStateRecord(snapshotState());
    sharedSyncMeta.updatedAt = normalizeStoredDate(payload.updatedAt || "", new Date().toISOString());
    workspaceLibrary.pages = payload.pages || workspaceLibrary.pages;
    syncCurrentWorkspaceFromLibrary();
    renderWorkspacePages();
    markSharedSyncOnline();
  } catch (error) {
    sharedSync.pendingPush = true;
    markSharedSyncOffline(error);
    console.error("Synchronisation BDD impossible", error);
  } finally {
    sharedSync.requestInFlight = false;
    if (sharedSync.online && sharedSync.pendingPush) {
      void flushSharedStateSyncQueue();
    }
  }
}

async function pullSharedStateFromServer() {
  if (!sharedStateStore || sharedSync.requestInFlight) {
    return;
  }

  sharedSync.requestInFlight = true;

  try {
    const record = await fetchSharedStateRecord();
    if (!record?.state) {
      return;
    }

    if (record.updatedAt && record.updatedAt === sharedSyncMeta.updatedAt) {
      workspaceLibrary.pages = record.pages || workspaceLibrary.pages;
      syncCurrentWorkspaceFromLibrary();
      renderWorkspacePages();
      return;
    }

    hydrateState(record.state);
    if (refreshStoredDeliveryNoteAnalyses()) {
      queueSharedStateSync();
    }
    sharedSyncMeta.updatedAt = record.updatedAt;
    workspaceLibrary.pages = record.pages || workspaceLibrary.pages;
    syncCurrentWorkspaceFromLibrary();
    render();
  } finally {
    sharedSync.requestInFlight = false;
  }
}

async function fetchSharedStateRecord() {
  if (!sharedStateStore) {
    return null;
  }

  try {
    const payload = await sharedStateStore.fetchStateRecord();
    markSharedSyncOnline();
    return {
      state: payload?.state || null,
      updatedAt: normalizeStoredDate(payload?.updatedAt || "", ""),
      pages: Array.isArray(payload?.pages) ? payload.pages : [],
    };
  } catch (error) {
    markSharedSyncOffline(error);
    console.error("Lecture BDD impossible", error);
    return null;
  }
}

function markSharedSyncOnline() {
  sharedSync.online = true;
  sharedSync.offlineToastShown = false;
  syncSharedStateBadge("shared");
}

function markSharedSyncOffline(error = null) {
  sharedSync.online = false;
  syncSharedStateBadge("offline");

  if (sharedSync.offlineToastShown) {
    return;
  }

  sharedSync.offlineToastShown = true;
  showToast(getSharedSyncErrorMessage(error), "danger");
}

function getSharedSyncErrorMessage(error) {
  const code = String(error?.code || "");
  const message = normalizeFreeText(String(error?.message || ""));

  if (code === "42P01" || /shared_state/i.test(message)) {
    return "La table Supabase n'existe pas encore. Relancez le script SQL puis rechargez la page.";
  }

  if (code === "42703" || /created_at|title/.test(message)) {
    return "La structure Supabase est incomplete. Relancez le script SQL de migration puis rechargez la page.";
  }

  if (/permission|row-level security|rls|42501/i.test(`${code} ${message}`)) {
    return "Supabase refuse l'acces aux donnees. Verifiez la connexion et les policies RLS.";
  }

  if (/invalid login credentials|invalid_credentials/i.test(message)) {
    return "Code d'acces incorrect.";
  }

  if (/email not confirmed/i.test(message)) {
    return "Le compte Supabase n'est pas confirme. Verifiez l'utilisateur d'acces dans Supabase Auth.";
  }

  if (/user not found|invalid/i.test(message) && /auth/i.test(String(error?.name || ""))) {
    return "Le compte d'acces Supabase est introuvable. Creez l'utilisateur de partage puis reessayez.";
  }

  if (/supabase-unavailable/i.test(message)) {
    return "Le client Supabase n'a pas pu etre charge. Rechargez la page puis reessayez.";
  }

  return "La BDD cloud ne repond plus. Aucun scan n'est sauvegarde tant que la connexion n'est pas revenue.";
}

function snapshotState() {
  if (typeof globalThis.structuredClone === "function") {
    return globalThis.structuredClone(state);
  }

  return JSON.parse(JSON.stringify(state));
}

function render() {
  renderHeroStats();
  renderBaqueSelect();
  renderWorkspacePages();
  renderSmallParcelScans();
  renderDestinationRuleTargetOptions();
  renderDestinationRules();
  renderDestinationSummary();
  safelyRenderSection(renderSortingPlan, renderSortingPlanFallback);
  safelyRenderSection(renderBaques, renderBaquesFallback);
  renderSearchResults();
  renderDeliveryNotes();
  applyCollapseStateToDom();
}

function renderWorkspacePages() {
  if (!ui.workspaceList) {
    return;
  }

  const pages = workspaceLibrary.pages.length
    ? workspaceLibrary.pages
    : [{
      id: workspacePage.id,
      title: workspacePage.title,
      createdAt: "",
      updatedAt: "",
      parcelsCount: state.parcels.length + getSmallParcelCountTotal(),
      baquesCount: state.baques.length,
    }];

  ui.workspaceList.innerHTML = pages
    .map((page) => {
      const isCurrent = page.id === workspacePage.id;
      const isPrimary = page.id === "global";
      const countBits = [
        `<span class="distribution-chip">${escapeHtml(String(page.parcelsCount || 0))} colis</span>`,
        `<span class="distribution-chip">${escapeHtml(String(page.baquesCount || 0))} baques</span>`,
      ];
      const statusBits = [
        isPrimary ? `<span class="distribution-chip">Principale</span>` : "",
        isCurrent ? `<span class="status-badge">Page actuelle</span>` : "",
      ].filter(Boolean);
      const managementButtons = [
        !isCurrent
          ? `<a class="btn btn--secondary document-card__action document-card__action--workspace" href="${escapeAttribute(buildWorkspacePageUrl(page.id, page.title))}">Ouvrir</a>`
          : "",
        !isPrimary
          ? `<button class="btn btn--secondary document-card__action document-card__action--workspace" type="button" data-workspace-action="rename-workspace" data-page-id="${escapeAttribute(page.id)}">Renommer</button>`
          : "",
        !isPrimary
          ? `<button class="btn btn--danger document-card__action document-card__action--workspace" type="button" data-workspace-action="delete-workspace" data-page-id="${escapeAttribute(page.id)}">Supprimer</button>`
          : "",
      ].filter(Boolean);
      const updatedAtLabel = page.updatedAt ? `Mise a jour ${escapeHtml(formatDate(page.updatedAt))}` : "";
      const createdAtLabel = !updatedAtLabel && page.createdAt ? `Creee le ${escapeHtml(formatDate(page.createdAt))}` : "";
      const timestampLabel = updatedAtLabel || createdAtLabel;

      return `
        <article class="document-card document-card--workspace${isCurrent ? " document-card--active" : ""}">
          <div class="document-card__topline">
            <p class="document-card__title">${escapeHtml(page.title || "Page sans nom")}</p>
            ${statusBits.length ? `<div class="document-card__badges">${statusBits.join("")}</div>` : ""}
          </div>
          <p class="document-card__id">ID : ${escapeHtml(page.id)}</p>
          <div class="document-summary document-summary--workspace">
            ${countBits.join("")}
          </div>
          <div class="document-card__footer document-card__footer--workspace">
            ${timestampLabel
              ? `<p class="document-card__meta document-card__meta--workspace">${timestampLabel}</p>`
              : `<span class="document-card__meta-spacer" aria-hidden="true"></span>`}
            ${managementButtons.length
              ? `<div class="document-card__actions document-card__actions--workspace">${managementButtons.join("")}</div>`
              : ""}
          </div>
        </article>
      `;
    })
    .join("");
}

function safelyRenderSection(renderSection, fallbackRender) {
  try {
    renderSection();
  } catch (error) {
    console.error("Erreur de rendu", renderSection?.name || "section", error);
    fallbackRender?.(error);
  }
}

function renderHeroStats() {
  const totalBaques = state.baques.length;
  const totalParcels = state.parcels.length + getSmallParcelCountTotal();
  const totalDestinations = getDestinationGroups().length;

  ui.heroStats.innerHTML = [
    statCard(totalBaques, "Baques actives"),
    statCard(totalParcels, "Colis suivis"),
    statCard(totalDestinations, "Destinations"),
  ].join("");
}

function statCard(value, label) {
  return `
    <article class="stat-card">
      <strong>${escapeHtml(String(value))}</strong>
      <span>${escapeHtml(label)}</span>
    </article>
  `;
}

function renderBaqueSelect() {
  const previousValue = ui.parcelBaqueSelect.value;
  const orderedBaques = getOrderedBaquesForLayout();

  ui.parcelBaqueSelect.innerHTML = orderedBaques
    .map(
      (baque) => `
        <option value="${escapeHtml(baque.id)}">
          ${escapeHtml(baque.name)} - ${escapeHtml(baque.location)}
        </option>
      `,
    )
    .join("");

  ui.parcelBaqueSelect.value = orderedBaques.some((baque) => baque.id === previousValue)
    ? previousValue
    : orderedBaques[0]?.id || "";
}

function renderDestinationSummary() {
  const destinationGroups = getDestinationGroups();

  if (!destinationGroups.length) {
    ui.destinationSummary.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucun colis pour le moment. La vue par destination apparaitra ici.</p>
      </article>
    `;
    return;
  }

  ui.destinationSummary.innerHTML = destinationGroups
    .map((group) => {
      const heading = getDestinationCardHeading(group);
      const chips = group.distribution
        .map(
          ([baqueName, count]) => `
            <span class="distribution-chip">${escapeHtml(baqueName)} : ${escapeHtml(String(count))}</span>
          `,
        )
        .join("");
      const destinations = group.rule && group.destinations.length
        ? `
          <div class="destination-card__aliases">
            ${group.destinations
              .map(
                ([destination, count]) => `
                  <span class="tag">${escapeHtml(destination)}${count > 1 ? ` x${escapeHtml(String(count))}` : ""}</span>
                `,
              )
              .join("")}
          </div>
        `
        : "";
      const quickAction = group.rule
        ? ""
        : `
          <button
            class="btn btn--secondary destination-card__action"
            type="button"
            data-action="prefill-destination-rule"
            data-label="${escapeAttribute(group.label)}"
          >
            Creer une regle
          </button>
        `;

      return `
        <article class="destination-card${group.rule ? " destination-card--rule" : ""}">
          <div class="destination-card__top">
            <div class="destination-card__title-block">
              ${heading.eyebrow ? `<p class="destination-card__eyebrow">${escapeHtml(heading.eyebrow)}</p>` : ""}
              <div class="destination-card__title-row">
                <h3>${escapeHtml(heading.title)}</h3>
                ${group.rule ? `<span class="tag">Regle</span>` : ""}
              </div>
              ${heading.subtitle ? `<p class="destination-card__subtitle">${escapeHtml(heading.subtitle)}</p>` : ""}
            </div>
            ${quickAction ? `<div class="destination-card__toolbar">${quickAction}</div>` : ""}
          </div>
          <div class="destination-count">${escapeHtml(String(group.parcels.length))}</div>
          <div class="destination-card__meta">
            <span>${escapeHtml(String(group.parcels.length))} colis</span>
            ${renderRouteCodeMeta(group.parcels)}
            ${group.rule ? `<span><strong>Condition :</strong> ${escapeHtml(getDestinationRuleMatchModeLabel(group.rule.matchMode))}</span>` : ""}
            ${group.rule?.preferredBaqueId ? `<span><strong>Baque cible :</strong> ${escapeHtml(getBaqueById(group.rule.preferredBaqueId)?.name || "Baque supprimee")}</span>` : ""}
          </div>
          ${destinations}
          <div class="distribution-list">${chips}</div>
        </article>
      `;
    })
    .join("");
}

function getDestinationCardHeading(group) {
  const label = normalizeFreeText(group?.label || "");
  if (!label) {
    return {
      eyebrow: group?.rule ? "Groupe" : "",
      title: "Sans destination",
      subtitle: "",
    };
  }

  if (group?.rule) {
    return {
      eyebrow: "Groupe",
      title: label,
      subtitle: "",
    };
  }

  const match = label.match(/^(\d{5})\s+(.+)$/u);
  if (!match) {
    return {
      eyebrow: "",
      title: label,
      subtitle: "",
    };
  }

  return {
    eyebrow: match[1],
    title: match[2],
    subtitle: "",
  };
}

function renderSortingPlan() {
  if (!ui.sortingPlan) {
    return;
  }

  const plans = getSortingPlans();
  if (!plans.length) {
    ui.sortingPlan.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Le plan de tri apparaitra ici des qu'une destination sera repartie sur plusieurs baques.</p>
      </article>
    `;
    return;
  }

  ui.sortingPlan.innerHTML = plans
    .map((plan, index) => {
      const effortLabel = getSortingEffortLabel(plan.totalEffort);
      const metrics = [
        `Baque cible : ${plan.targetBaque.name}`,
        `${plan.movedCount} colis a deplacer`,
        `${formatKnownWeightSummary(plan.totalMoveWeightKg, plan.unknownWeightCount)}`,
        `${plan.alreadyCount} deja sur place`,
      ];

      return `
        <article class="sorting-plan-card">
          <div class="sorting-plan-card__top">
            <div>
              <p class="section-kicker">Ordre ${escapeHtml(String(index + 1))}</p>
              <h4>${escapeHtml(plan.label)}</h4>
            </div>
            <span class="tag">${escapeHtml(effortLabel)}</span>
          </div>
          <div class="document-summary">
            ${metrics.map((metric) => `<span class="distribution-chip">${escapeHtml(metric)}</span>`).join("")}
          </div>
          <p class="field-help">
            Garder cette destination dans <strong>${escapeHtml(plan.targetBaque.name)}</strong> :
            ${escapeHtml(plan.alreadyCount)} colis y sont deja et c'est la position qui demande le moins d'effort cumule.
          </p>
          <div class="sorting-pass-list">
            ${plan.passes.map((pass, passIndex) => `
              <article class="sorting-pass">
                <div class="sorting-pass__head">
                  <strong>Passage ${escapeHtml(String(passIndex + 1))}</strong>
                  <span>${escapeHtml(pass.routeLabel)}</span>
                </div>
                <div class="sorting-pass__meta">
                  <span>${escapeHtml(String(pass.movedCount))} colis</span>
                  <span>${escapeHtml(formatKnownWeightSummary(pass.totalWeightKg, pass.unknownWeightCount))}</span>
                  <span>${escapeHtml(pass.advice)}</span>
                </div>
              </article>
            `).join("")}
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSortingPlanFallback() {
  if (!ui.sortingPlan) {
    return;
  }

  const fallbackPlans = buildFallbackSortingPlans();
  if (!fallbackPlans.length) {
    ui.sortingPlan.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Le plan de tri apparaitra ici des qu'une destination sera repartie sur plusieurs baques.</p>
      </article>
    `;
    return;
  }

  ui.sortingPlan.innerHTML = fallbackPlans
    .map((plan) => `
      <article class="sorting-plan-card">
        <div class="sorting-plan-card__top">
          <div>
            <p class="section-kicker">Plan simplifie</p>
            <h4>${escapeHtml(plan.label)}</h4>
          </div>
          <span class="tag">Repli</span>
        </div>
        <div class="document-summary">
          <span class="distribution-chip">Baque cible : ${escapeHtml(plan.targetBaqueName)}</span>
          <span class="distribution-chip">${escapeHtml(String(plan.movedCount))} colis a deplacer</span>
        </div>
        <p class="field-help">${escapeHtml(plan.routeLabel)}</p>
      </article>
    `)
    .join("");
}

function buildFallbackSortingPlans() {
  return getDestinationGroups()
    .map((group) => {
      const distribution = group.distribution || [];
      if (distribution.length < 2) {
        return null;
      }

      const target = distribution[0];
      if (!target) {
        return null;
      }

      return {
        label: group.label,
        targetBaqueName: target[0],
        movedCount: group.parcels.length - target[1],
        routeLabel: [...distribution.map(([baqueName]) => baqueName), target[0]].join(" -> "),
      };
    })
    .filter((plan) => plan && plan.movedCount > 0);
}

function getSortingPlans() {
  const destinationGroups = getDestinationGroups();
  const rawPlans = destinationGroups
    .map((group) => buildSortingPlan(group))
    .filter((plan) => plan && plan.movedCount > 0);

  return orderSortingPlans(rawPlans);
}

function buildSortingPlan(group) {
  const orderedBaques = getOrderedBaquesForLayout();
  if (orderedBaques.length < 2) {
    return null;
  }

  const layoutIndexByBaqueId = new Map(orderedBaques.map((baque, index) => [baque.id, index + 1]));
  const preferredBaque = group.rule?.preferredBaqueId ? getBaqueById(group.rule.preferredBaqueId) : null;
  const candidateBaques = preferredBaque ? [preferredBaque] : orderedBaques;
  const targetCandidates = candidateBaques
    .map((baque) => scoreSortingTargetBaque(group.parcels, baque, layoutIndexByBaqueId, orderedBaques.length))
    .filter(Boolean)
    .sort((left, right) => {
      const effortDiff = left.totalEffort - right.totalEffort;
      if (effortDiff !== 0) {
        return effortDiff;
      }

      const distanceDiff = left.totalDistance - right.totalDistance;
      if (distanceDiff !== 0) {
        return distanceDiff;
      }

      const movedDiff = left.movedCount - right.movedCount;
      if (movedDiff !== 0) {
        return movedDiff;
      }

      const alreadyDiff = right.alreadyCount - left.alreadyCount;
      if (alreadyDiff !== 0) {
        return alreadyDiff;
      }

      return left.targetPosition - right.targetPosition;
    });

  const bestTarget = targetCandidates[0];
  if (!bestTarget || !bestTarget.movedCount) {
    return null;
  }

  return {
    key: group.key,
    label: group.label,
    targetBaque: bestTarget.targetBaque,
    targetPosition: bestTarget.targetPosition,
    totalEffort: bestTarget.totalEffort,
    totalDistance: bestTarget.totalDistance,
    alreadyCount: bestTarget.alreadyCount,
    movedCount: bestTarget.movedCount,
    totalMoveWeightKg: bestTarget.totalMoveWeightKg,
    unknownWeightCount: bestTarget.unknownWeightCount,
    passes: buildSortingPasses(bestTarget),
  };
}

function scoreSortingTargetBaque(parcels, targetBaque, layoutIndexByBaqueId, totalBaques) {
  const targetPosition = layoutIndexByBaqueId.get(targetBaque.id);
  if (!targetPosition) {
    return null;
  }

  const sourceBucketsByBaqueId = new Map();
  let totalEffort = 0;
  let totalDistance = 0;
  let alreadyCount = 0;
  let movedCount = 0;
  let totalMoveWeightKg = 0;
  let unknownWeightCount = 0;

  parcels.forEach((parcel) => {
    const sourcePosition = layoutIndexByBaqueId.get(parcel.currentBaqueId);
    if (!sourcePosition) {
      return;
    }

    const distance = Math.abs(sourcePosition - targetPosition);
    const weightKg = parseParcelWeightKg(parcel);
    const handlingFactor = getParcelHandlingFactor(weightKg);

    totalEffort += distance * handlingFactor;
    totalDistance += distance;

    if (distance === 0) {
      alreadyCount += 1;
      return;
    }

    movedCount += 1;
    if (weightKg !== null) {
      totalMoveWeightKg += weightKg;
    } else {
      unknownWeightCount += 1;
    }

    const sourceBaque = getBaqueById(parcel.currentBaqueId);
    const bucket = sourceBucketsByBaqueId.get(parcel.currentBaqueId) || {
      baqueId: parcel.currentBaqueId,
      baqueName: sourceBaque?.name || "Baque",
      position: sourcePosition,
      movedCount: 0,
      totalWeightKg: 0,
      unknownWeightCount: 0,
      heaviestParcelKg: 0,
      effort: 0,
    };

    bucket.movedCount += 1;
    bucket.effort += distance * handlingFactor;
    if (weightKg !== null) {
      bucket.totalWeightKg += weightKg;
      bucket.heaviestParcelKg = Math.max(bucket.heaviestParcelKg, weightKg);
    } else {
      bucket.unknownWeightCount += 1;
    }

    sourceBucketsByBaqueId.set(parcel.currentBaqueId, bucket);
  });

  if (!movedCount) {
    return {
      targetBaque,
      targetPosition,
      totalEffort: 0,
      totalDistance: 0,
      alreadyCount,
      movedCount: 0,
      totalMoveWeightKg: 0,
      unknownWeightCount: 0,
      sourceBuckets: [],
    };
  }

  const centerPosition = (totalBaques + 1) / 2;
  if (totalMoveWeightKg >= 80) {
    totalEffort += Math.abs(targetPosition - centerPosition) * 0.35;
  }

  return {
    targetBaque,
    targetPosition,
    totalEffort,
    totalDistance,
    alreadyCount,
    movedCount,
    totalMoveWeightKg,
    unknownWeightCount,
    sourceBuckets: [...sourceBucketsByBaqueId.values()],
  };
}

function buildSortingPasses(plan) {
  const lowerSide = plan.sourceBuckets
    .filter((bucket) => bucket.position < plan.targetPosition)
    .sort((left, right) => left.position - right.position);
  const upperSide = plan.sourceBuckets
    .filter((bucket) => bucket.position > plan.targetPosition)
    .sort((left, right) => right.position - left.position);
  const lowerSummary = summarizeSortingSide(lowerSide);
  const upperSummary = summarizeSortingSide(upperSide);
  const sideOrder = [];

  if (lowerSide.length && upperSide.length) {
    if (upperSummary.effort > lowerSummary.effort) {
      sideOrder.push(upperSide, lowerSide);
    } else {
      sideOrder.push(lowerSide, upperSide);
    }
  } else if (lowerSide.length) {
    sideOrder.push(lowerSide);
  } else if (upperSide.length) {
    sideOrder.push(upperSide);
  }

  return sideOrder.map((sideBuckets) => {
    const summary = summarizeSortingSide(sideBuckets);
    const routeNames = [...sideBuckets.map((bucket) => bucket.baqueName), plan.targetBaque.name];

    return {
      routeLabel: routeNames.join(" -> "),
      movedCount: summary.movedCount,
      totalWeightKg: summary.totalWeightKg,
      unknownWeightCount: summary.unknownWeightCount,
      advice: getSortingHandlingAdvice(summary),
    };
  });
}

function summarizeSortingSide(sideBuckets) {
  return sideBuckets.reduce((summary, bucket) => ({
    movedCount: summary.movedCount + bucket.movedCount,
    totalWeightKg: summary.totalWeightKg + bucket.totalWeightKg,
    unknownWeightCount: summary.unknownWeightCount + bucket.unknownWeightCount,
    heaviestParcelKg: Math.max(summary.heaviestParcelKg, bucket.heaviestParcelKg),
    effort: summary.effort + bucket.effort,
  }), {
    movedCount: 0,
    totalWeightKg: 0,
    unknownWeightCount: 0,
    heaviestParcelKg: 0,
    effort: 0,
  });
}

function orderSortingPlans(plans) {
  const remaining = [...plans];
  const orderedPlans = [];
  let currentPosition = 1;

  while (remaining.length) {
    remaining.sort((left, right) => {
      const travelDiff = Math.abs(left.targetPosition - currentPosition) - Math.abs(right.targetPosition - currentPosition);
      if (travelDiff !== 0) {
        return travelDiff;
      }

      const weightDiff = right.totalMoveWeightKg - left.totalMoveWeightKg;
      if (weightDiff !== 0) {
        return weightDiff;
      }

      const movedDiff = right.movedCount - left.movedCount;
      if (movedDiff !== 0) {
        return movedDiff;
      }

      return left.totalEffort - right.totalEffort;
    });

    const nextPlan = remaining.shift();
    orderedPlans.push(nextPlan);
    currentPosition = nextPlan.targetPosition;
  }

  return orderedPlans;
}

function renderDestinationRules() {
  if (!state.destinationRules.length) {
    ui.destinationRulesList.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucune regle pour le moment. Ajoutez-en une pour regrouper des destinations proches.</p>
      </article>
    `;
    return;
  }

  ui.destinationRulesList.innerHTML = getSortedDestinationRules()
    .map((rule) => {
      const matchedDestinations = getRuleMatchedDestinationLabels(rule);
      const matchedParcelsCount = countRuleMatchedParcels(rule);
      const previewTags = matchedDestinations.length
        ? matchedDestinations
          .slice(0, 4)
          .map(
            (destination) => `
              <span class="tag">${escapeHtml(destination)}</span>
            `,
          )
          .join("")
        : `<span class="tag">Aucun colis correspondant</span>`;
      const extraCount = matchedDestinations.length > 4
        ? `<span class="tag">+${escapeHtml(String(matchedDestinations.length - 4))}</span>`
        : "";

      return `
        <article class="destination-rule-card" data-rule-id="${escapeHtml(rule.id)}">
          <div class="destination-rule-card__top">
            <span class="count-pill">${escapeHtml(String(matchedParcelsCount))} colis</span>
            <button class="btn btn--danger" type="button" data-action="delete-destination-rule" data-rule-id="${escapeHtml(rule.id)}">
              Supprimer
            </button>
          </div>
          <div class="field-grid destination-rule-card__grid">
            <label class="field">
              <span>Nom du groupe</span>
              <input
                type="text"
                value="${escapeAttribute(rule.label)}"
                data-field="label"
                data-rule-id="${escapeHtml(rule.id)}"
                aria-label="Nom du groupe"
              >
            </label>

            <label class="field">
              <span>Condition</span>
              <select data-field="matchMode" data-rule-id="${escapeHtml(rule.id)}" aria-label="Condition de regroupement">
                <option value="any" ${rule.matchMode === "any" ? "selected" : ""}>Au moins un mot-cle</option>
                <option value="all" ${rule.matchMode === "all" ? "selected" : ""}>Tous les mots-cles</option>
              </select>
            </label>

            <label class="field">
              <span>Baque cible</span>
              <select data-field="preferredBaqueId" data-rule-id="${escapeHtml(rule.id)}" aria-label="Baque cible">
                ${buildBaqueTargetOptions(rule.preferredBaqueId || "")}
              </select>
            </label>
          </div>

          <label class="field">
            <span>Mots-cles adresse</span>
            <textarea
              rows="3"
              data-field="patterns"
              data-rule-id="${escapeHtml(rule.id)}"
              aria-label="Mots-cles adresse"
            >${escapeHtml(rule.patterns.join("\n"))}</textarea>
          </label>

          <div class="destination-rule-card__matches">
            ${previewTags}
            ${extraCount}
          </div>
        </article>
      `;
    })
    .join("");
}

function renderDestinationRuleTargetOptions() {
  if (!ui.destinationRuleTargetBaqueSelect) {
    return;
  }

  const previousValue = ui.destinationRuleTargetBaqueSelect.value;
  ui.destinationRuleTargetBaqueSelect.innerHTML = buildBaqueTargetOptions(previousValue);
  ui.destinationRuleTargetBaqueSelect.value = hasBaqueId(previousValue) ? previousValue : "";
}

function buildBaqueTargetOptions(selectedBaqueId = "") {
  return [
    `<option value="" ${!selectedBaqueId ? "selected" : ""}>Calcul automatique</option>`,
    ...getOrderedBaquesForLayout().map(
      (baque) => `
        <option value="${escapeHtml(baque.id)}" ${baque.id === selectedBaqueId ? "selected" : ""}>
          ${escapeHtml(baque.name)}
        </option>
      `,
    ),
  ].join("");
}

function renderBaques() {
  if (!ui.baquesGrid) {
    return;
  }

  const baques = getOrderedBaquesForLayout();
  if (!baques.length) {
    ui.baquesGrid.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucune baque disponible. Rechargez la page pour restaurer les emplacements par defaut.</p>
      </article>
    `;
    return;
  }

  ui.baquesGrid.innerHTML = baques
    .map((baque) => renderBaqueCard(baque))
    .join("");
}

function renderBaquesFallback() {
  if (!ui.baquesGrid) {
    return;
  }

  const baques = getOrderedBaquesForLayout();
  if (!baques.length) {
    ui.baquesGrid.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucune baque disponible.</p>
      </article>
    `;
    return;
  }

  ui.baquesGrid.innerHTML = baques
    .map((baque) => {
      const parcels = getParcelsForBaque(baque.id);
      return `
        <article class="baque-card" data-baque-id="${escapeHtml(baque.id)}">
          <div class="baque-card__top">
            <div class="baque-card__status">
              <span class="count-pill">${escapeHtml(String(parcels.length))} colis</span>
            </div>
            <p class="parcel-code">${escapeHtml(baque.name)}</p>
            <p class="parcel-meta">${escapeHtml(baque.location)}</p>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderBaqueCard(baque) {
  const parcels = getParcelsForBaque(baque.id);
  const validationBadge = baque.validatedAt
    ? `<span class="validation-pill">Validee</span>`
    : "";
  const validationMeta = baque.validatedAt
    ? `<p class="baque-card__validation-meta">Validee le ${escapeHtml(formatDate(baque.validatedAt))}</p>`
    : `<p class="baque-card__validation-meta">Baque non validee</p>`;
  const parcelMarkup = parcels.length
    ? parcels.map((parcel) => safeParcelTemplate(parcel)).join("")
    : emptyBaqueTemplate();

  return `
    <article class="baque-card${baque.validatedAt ? " baque-card--validated" : ""}" data-baque-id="${escapeHtml(baque.id)}">
      <div class="baque-card__top">
        <div class="baque-card__meta">
          <div class="baque-card__status">
            <span class="count-pill">${escapeHtml(String(parcels.length))} ${escapeHtml(pluralize(parcels.length, "colis", "colis"))}</span>
            ${validationBadge}
          </div>
          <div class="baque-card__actions">
            <button class="btn btn--secondary" type="button" data-action="toggle-baque-validation" data-baque-id="${escapeHtml(baque.id)}">
              ${baque.validatedAt ? "Retirer la validation" : "Valider la baque"}
            </button>
            <button class="btn btn--danger" type="button" data-action="delete-baque" data-baque-id="${escapeHtml(baque.id)}">
              Supprimer la baque
            </button>
          </div>
        </div>

        <label class="baque-card__title">
          <input
            type="text"
            value="${escapeAttribute(baque.name)}"
            data-field="name"
            data-baque-id="${escapeHtml(baque.id)}"
            aria-label="Nom de la baque"
          >
        </label>

        <label class="baque-card__location">
          <input
            type="text"
            value="${escapeAttribute(baque.location)}"
            data-field="location"
            data-baque-id="${escapeHtml(baque.id)}"
            aria-label="Emplacement de la baque"
          >
        </label>
        ${validationMeta}
      </div>

      <div class="parcel-list">
        ${parcelMarkup}
      </div>
    </article>
  `;
}

function safeParcelTemplate(parcel) {
  try {
    return parcelTemplate(parcel);
  } catch (error) {
    console.error("Rendu colis impossible", parcel, error);
    return `
      <article class="parcel-item">
        <div class="parcel-item__top">
          <div>
            <p class="parcel-code">${escapeHtml(getParcelIdentifier(parcel))}</p>
            <p class="parcel-meta">Ce colis contient une donnee invalide. Ouvrez-le depuis un autre appareil ou rescanez-le.</p>
          </div>
        </div>
      </article>
    `;
  }
}

function parcelTemplate(parcel) {
  const options = getOrderedBaquesForLayout()
    .map(
      (baque) => `
        <option value="${escapeHtml(baque.id)}" ${baque.id === parcel.currentBaqueId ? "selected" : ""}>
          ${escapeHtml(baque.name)}
        </option>
      `,
    )
    .join("");
  const displayDestination = getParcelDestinationDisplay(parcel);
  const displayRouteCode = formatRouteCodeForDisplay(parcel.routeCode);

  const primaryLines = [
    displayDestination ? `Destination : <strong>${escapeHtml(displayDestination)}</strong>` : "",
    parcel.client ? `Client : ${escapeHtml(parcel.client)}` : "",
    parcel.packageIndex ? `Colis : ${escapeHtml(parcel.packageIndex)}` : "",
    parcel.weight ? `Poids : ${escapeHtml(parcel.weight)}` : "",
    parcel.measuredDimensions ? `Mesures reelles : ${escapeHtml(parcel.measuredDimensions)}` : "",
  ]
    .filter(Boolean)
    .join("<br>");

  const detailLines = [
    parcel.commandNumber ? `Numero de commande : ${escapeHtml(parcel.commandNumber)}` : "",
    parcel.barcode && parcel.barcode !== parcel.commandNumber ? `Code-barres : ${escapeHtml(parcel.barcode)}` : "",
    parcel.routeCode ? `Numero destination : ${escapeHtml(displayRouteCode)}` : "",
    parcel.routeLabel ? `Route : ${escapeHtml(parcel.routeLabel)}` : "",
    parcel.reference ? `Reference : ${escapeHtml(parcel.reference)}` : "",
    parcel.shippingDate ? `Date : ${escapeHtml(parcel.shippingDate)}` : "",
    `Origine : ${escapeHtml(getOriginLabel(parcel))}`,
    `Derniere mise a jour : ${escapeHtml(formatDate(parcel.updatedAt || parcel.createdAt))}`,
  ]
    .filter(Boolean)
    .join("<br>");

  const detailsMarkup = detailLines
    ? `
      <details class="parcel-details">
        <summary>Plus d'infos</summary>
        <p class="parcel-meta parcel-meta--details">${detailLines}</p>
      </details>
    `
    : "";
  const tagLabel = displayRouteCode || getDestinationShortLabel(displayDestination) || "Colis";
  const parcelHeading = getParcelIdentifier(parcel);

  return `
    <article class="parcel-item" data-parcel-id="${escapeHtml(parcel.id)}">
      <div class="parcel-item__top">
        <div>
          <p class="parcel-code">${escapeHtml(parcelHeading)}</p>
          <p class="parcel-meta">${primaryLines}</p>
          ${detailsMarkup}
        </div>
        <span class="tag">${escapeHtml(tagLabel)}</span>
      </div>

      <div class="parcel-item__bottom">
        <div class="move-group">
          <select data-role="move-select" data-parcel-id="${escapeHtml(parcel.id)}" aria-label="Choisir une autre baque">
            ${options}
          </select>
          <button class="btn btn--secondary" type="button" data-action="move-parcel" data-parcel-id="${escapeHtml(parcel.id)}">
            Deplacer
          </button>
        </div>

        <button class="btn btn--danger" type="button" data-action="delete-parcel" data-parcel-id="${escapeHtml(parcel.id)}">
          Supprimer
        </button>
      </div>
    </article>
  `;
}

function emptyBaqueTemplate() {
  return `
    <div class="empty-state">
      Aucun colis dans cette baque.
    </div>
  `;
}

function getSmallParcelCountTotal(scans = state.smallParcelScans) {
  return scans.reduce((total, scan) => total + Math.max(1, Number(scan.quantity || 1)), 0);
}

function expandSmallParcelScansForComparison() {
  return state.smallParcelScans.flatMap((scan) => {
    const quantity = Math.max(1, Number(scan.quantity || 1));
    return Array.from({ length: quantity }, () => ({
      commandNumber: scan.commandNumber,
      barcode: scan.barcode || scan.commandNumber,
      routeCode: "",
      destination: "",
      client: "",
      description: "",
      routeLabel: "",
      reference: "",
      shippingDate: "",
      weight: "",
      measuredDimensions: "",
      packageIndex: "",
    }));
  });
}

function getPdfComparableParcels() {
  return [...state.parcels, ...expandSmallParcelScansForComparison()];
}

function normalizeComparableDeliveryEntries(entries) {
  if (!Array.isArray(entries)) {
    return [];
  }

  return entries
    .map((entry) => ({
      commandNumber: normalizeCommandNumber(entry?.commandNumber || ""),
      expectedCount: Math.max(1, Number(entry?.expectedCount || 1)),
      client: normalizeFreeText(entry?.client || ""),
      city: normalizeFreeText(entry?.city || ""),
      rawContext: normalizeFreeText(entry?.rawContext || ""),
    }))
    .filter((entry) => entry.commandNumber);
}

function buildStoredDeliveryNoteAnalysis(entries, analyzedAt = "") {
  const normalizedEntries = normalizeComparableDeliveryEntries(entries);
  const analysis = compareDeliveryNoteEntries(normalizedEntries, getPdfComparableParcels());
  return {
    ...analysis,
    entries: normalizedEntries,
    analyzedAt: normalizeStoredDate(analyzedAt || "", new Date().toISOString()),
  };
}

function refreshStoredDeliveryNoteAnalyses() {
  const now = new Date().toISOString();
  let updated = false;

  state.deliveryNotes.forEach((note) => {
    if (!Array.isArray(note?.analysis?.entries) || !note.analysis.entries.length) {
      return;
    }

    const nextAnalysis = buildStoredDeliveryNoteAnalysis(note.analysis.entries, note.analysis.analyzedAt);
    if (JSON.stringify(note.analysis) === JSON.stringify(nextAnalysis)) {
      return;
    }

    note.analysis = nextAnalysis;
    note.updatedAt = now;
    updated = true;
  });

  return updated;
}

function renderSmallParcelScans() {
  if (!ui.smallParcelList || !ui.smallParcelSummary) {
    return;
  }

  const totalCount = getSmallParcelCountTotal();
  const distinctCommands = new Set(state.smallParcelScans.map((scan) => scan.commandNumber)).size;
  ui.smallParcelSummary.innerHTML = totalCount
    ? `
      <span class="distribution-chip">${escapeHtml(String(totalCount))} ${escapeHtml(pluralize(totalCount, "petit colis compte", "petits colis comptes"))}</span>
      <span class="distribution-chip">${escapeHtml(String(distinctCommands))} ${escapeHtml(pluralize(distinctCommands, "commande", "commandes"))}</span>
    `
    : "";

  if (!state.smallParcelScans.length) {
    ui.smallParcelList.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucun petit colis compte pour le moment.</p>
      </article>
    `;
    return;
  }

  ui.smallParcelList.innerHTML = state.smallParcelScans
    .slice()
    .sort((left, right) => new Date(right.updatedAt) - new Date(left.updatedAt))
    .map((scan) => {
      const quantity = Math.max(1, Number(scan.quantity || 1));
      const metaBits = [
        `Ajoute le ${escapeHtml(formatDate(scan.updatedAt || scan.createdAt))}`,
        scan.barcode && scan.barcode !== scan.commandNumber ? `Code lu : ${escapeHtml(scan.barcode)}` : "",
      ].filter(Boolean);

      return `
        <article class="document-card">
          <div class="document-card__body">
            <p class="document-card__title">Commande ${escapeHtml(scan.commandNumber)}</p>
            <div class="document-summary">
              <span class="distribution-chip">${escapeHtml(String(quantity))} ${escapeHtml(pluralize(quantity, "petit colis", "petits colis"))}</span>
            </div>
            <p class="document-card__meta">${metaBits.join(" • ")}</p>
          </div>
          <div class="document-card__actions">
            <button
              class="btn btn--danger document-card__action"
              type="button"
              data-action="delete-small-parcel-scan"
              data-small-scan-id="${escapeAttribute(scan.id)}"
            >
              Supprimer
            </button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSearchResults() {
  const query = ui.searchInput.value.trim().toLowerCase();

  if (!query) {
    ui.searchResults.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Saisissez un code-barres ou une destination pour retrouver un colis.</p>
      </article>
    `;
    return;
  }

  const matches = state.parcels.filter((parcel) => {
    const baque = getBaqueById(parcel.currentBaqueId);
    const haystack = [
      getParcelCommandNumber(parcel),
      parcel.barcode,
      parcel.routeCode || "",
      getParcelDestinationDisplay(parcel),
      parcel.client || "",
      parcel.description || "",
      parcel.routeLabel || "",
      parcel.reference || "",
      parcel.shippingDate || "",
      parcel.weight || "",
      parcel.measuredDimensions || "",
      parcel.packageIndex || "",
      baque?.name || "",
      baque?.location || "",
      getOriginLabel(parcel),
    ]
      .join(" ")
      .toLowerCase();

    return haystack.includes(query);
  });

  if (!matches.length) {
    ui.searchResults.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucun colis ne correspond a votre recherche.</p>
      </article>
    `;
    return;
  }

  ui.searchResults.innerHTML = matches
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
    .map((parcel) => {
      const baque = getBaqueById(parcel.currentBaqueId);
      const displayDestination = getParcelDestinationDisplay(parcel);

      return `
        <article class="search-card">
          <h3>${escapeHtml(getParcelIdentifier(parcel))}</h3>
          <div class="search-card__meta">
            ${parcel.commandNumber ? `<span><strong>Numero de commande :</strong> ${escapeHtml(parcel.commandNumber)}</span>` : ""}
            ${parcel.barcode && parcel.barcode !== parcel.commandNumber ? `<span><strong>Code-barres :</strong> ${escapeHtml(parcel.barcode)}</span>` : ""}
            ${parcel.routeCode ? `<span><strong>Numero destination :</strong> ${escapeHtml(formatRouteCodeForDisplay(parcel.routeCode))}</span>` : ""}
            <span><strong>Destination :</strong> ${escapeHtml(displayDestination)}</span>
            ${parcel.client ? `<span><strong>Client :</strong> ${escapeHtml(parcel.client)}</span>` : ""}
            ${parcel.description ? `<span><strong>Description :</strong> ${escapeHtml(parcel.description)}</span>` : ""}
            ${parcel.routeLabel ? `<span><strong>Route :</strong> ${escapeHtml(parcel.routeLabel)}</span>` : ""}
            ${parcel.reference ? `<span><strong>Reference :</strong> ${escapeHtml(parcel.reference)}</span>` : ""}
            ${parcel.packageIndex ? `<span><strong>Colis :</strong> ${escapeHtml(parcel.packageIndex)}</span>` : ""}
            ${parcel.weight ? `<span><strong>Poids :</strong> ${escapeHtml(parcel.weight)}</span>` : ""}
            ${parcel.measuredDimensions ? `<span><strong>Mesures reelles :</strong> ${escapeHtml(parcel.measuredDimensions)}</span>` : ""}
            ${parcel.shippingDate ? `<span><strong>Date :</strong> ${escapeHtml(parcel.shippingDate)}</span>` : ""}
            <span><strong>Baque actuelle :</strong> ${escapeHtml(baque?.name || "Baque supprimee")}</span>
            <span><strong>Emplacement :</strong> ${escapeHtml(baque?.location || "Inconnu")}</span>
            <span><strong>Origine :</strong> ${escapeHtml(getOriginLabel(parcel))}</span>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderDeliveryNotes() {
  if (!state.deliveryNotes.length) {
    ui.deliveryNoteList.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucun PDF importe pour le moment.</p>
      </article>
    `;
    return;
  }

  ui.deliveryNoteList.innerHTML = state.deliveryNotes
    .slice()
    .sort((left, right) => new Date(right.importedAt) - new Date(left.importedAt))
    .map((note) => deliveryNoteTemplate(note))
    .join("");
}

function deliveryNoteTemplate(note) {
  const analysis = note.analysis || null;
  const zeroMatchNote = analysis && !analysis.parseError && analysis.totalEntries > 0 && analysis.totalRegisteredCount === 0
    ? `<p class="field-help">${escapeHtml("Aucune correspondance detectee avec les colis enregistres. Verifiez que ce bon de livraison correspond bien a la meme tournee et que les colis ont un numero de commande saisi ou scanne.")}</p>`
    : "";
  const incomparablesNote = analysis?.incomparableParcelsCount
    ? `<p class="field-help">${escapeHtml(`${analysis.incomparableParcelsCount} colis enregistres sans numero de commande ne peuvent pas etre compares au PDF.`)}</p>`
    : "";
  const resultDetails = analysis?.parseError
    ? `<p class="field-help">${escapeHtml(analysis.parseError)}</p>`
    : analysis?.missingEntries.length
    ? `
      <div class="document-missing">
        <p class="document-missing__title">Colis manquants</p>
        <div class="document-missing__list">
          ${analysis.missingEntries.map((entry) => `
            <article class="document-missing__item">
              <strong>${escapeHtml(entry.commandNumber)}</strong>
              <span>${escapeHtml(String(entry.registeredCount))} / ${escapeHtml(String(entry.expectedCount))} enregistres</span>
              <span>${escapeHtml(String(entry.missingCount))} colis manquant${entry.missingCount > 1 ? "s" : ""}</span>
              ${entry.client ? `<span>${escapeHtml(entry.client)}</span>` : ""}
              ${entry.city ? `<span>${escapeHtml(entry.city)}</span>` : ""}
            </article>
          `).join("")}
        </div>
      </div>
    `
    : `<p class="field-help">Toutes les livraisons du PDF semblent deja enregistrees.</p>`;
  const summary = analysis
    ? `
      <div class="document-summary">
        <span class="distribution-chip">Lignes : ${escapeHtml(String(analysis.totalEntries))}</span>
        <span class="distribution-chip">Attendus : ${escapeHtml(String(analysis.totalExpectedCount))}</span>
        <span class="distribution-chip">Enregistres : ${escapeHtml(String(analysis.totalRegisteredCount))}</span>
        <span class="distribution-chip distribution-chip--alert">Manquants : ${escapeHtml(String(analysis.totalMissingCount))}</span>
      </div>
      ${zeroMatchNote}
      ${incomparablesNote}
      ${resultDetails}
    `
    : `
      <p class="field-help">Aucune analyse de comparaison disponible pour ce PDF.</p>
    `;

  return `
    <article class="document-card">
      <div class="document-card__body">
        <p class="document-card__title">${escapeHtml(note.name)}</p>
        <p class="document-card__meta">
          PDF ${escapeHtml(formatFileSize(note.size))} | importe le ${escapeHtml(formatDate(note.importedAt))}
          ${analysis?.analyzedAt ? ` | analyse le ${escapeHtml(formatDate(analysis.analyzedAt))}` : ""}
        </p>
        ${summary}
      </div>
      <div class="document-card__actions">
        <button
          class="btn btn--secondary document-card__action"
          type="button"
          data-action="analyze-delivery-note"
          data-note-id="${escapeHtml(note.id)}"
        >
          ${analysis ? "Reanalyser" : "Analyser"}
        </button>
        <button
          class="btn btn--secondary document-card__action"
          type="button"
          data-action="simulate-delivery-note"
          data-note-id="${escapeHtml(note.id)}"
        >
          Simuler les colis
        </button>
        <button
          class="btn btn--danger document-card__action"
          type="button"
          data-action="delete-delivery-note"
          data-note-id="${escapeHtml(note.id)}"
        >
          Supprimer
        </button>
      </div>
    </article>
  `;
}

function handleParcelSubmit(event) {
  event.preventDefault();
  upsertParcel();
}

function handleSmallParcelSubmit(event) {
  event.preventDefault();
  upsertSmallParcelScan();
}

function handleSmallParcelListClick(event) {
  const button = event.target.closest('[data-action="delete-small-parcel-scan"]');
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const scanId = button.dataset.smallScanId || "";
  if (!scanId) {
    return;
  }

  deleteSmallParcelScan(scanId);
}

function handleBaqueSubmit(event) {
  event.preventDefault();

  const name = ui.baqueNameInput.value.trim();
  const location = ui.baqueLocationInput.value.trim();

  if (!name || !location) {
    showToast("Le nom et l'emplacement de la baque sont obligatoires.", "danger");
    return;
  }

  const now = new Date().toISOString();
  state.baques.push({
    id: createId(),
    name,
    location,
    validatedAt: "",
    createdAt: now,
    updatedAt: now,
  });

  saveState();
  render();
  ui.baqueForm.reset();
  showToast(`La baque "${name}" a ete ajoutee.`);
}

function handleDestinationRuleSubmit(event) {
  event.preventDefault();

  const label = normalizeFreeText(ui.destinationRuleLabelInput.value);
  const matchMode = normalizeDestinationRuleMatchMode(ui.destinationRuleMatchModeSelect.value);
  const preferredBaqueId = normalizeDestinationRuleTargetBaqueId(ui.destinationRuleTargetBaqueSelect.value);
  const patterns = parseDestinationRulePatterns(ui.destinationRulePatternsInput.value);

  if (!label || !patterns.length) {
    showToast("Le nom du groupe et au moins un mot-cle sont obligatoires.", "danger");
    return;
  }

  const now = new Date().toISOString();
  state.destinationRules.push({
    id: createId(),
    label,
    matchMode,
    preferredBaqueId,
    patterns,
    createdAt: now,
    updatedAt: now,
  });

  saveState();
  render();
  ui.destinationRuleForm.reset();
  ui.destinationRuleMatchModeSelect.value = "any";
  ui.destinationRuleTargetBaqueSelect.value = "";
  showToast(`Regle "${label}" ajoutee.`);
}

function handleDestinationSummaryClick(event) {
  const button = event.target.closest('[data-action="prefill-destination-rule"]');
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const label = normalizeFreeText(button.dataset.label || "");
  if (!label) {
    return;
  }

  ui.destinationRuleLabelInput.value = label;
  ui.destinationRulePatternsInput.value = label;
  ui.destinationRuleMatchModeSelect.value = "any";
  ui.destinationRuleTargetBaqueSelect.value = "";
  ui.destinationRuleLabelInput.focus();
  showToast("Regle pre-remplie. Ajustez-la puis enregistrez.");
}

function handleDestinationRulesClick(event) {
  const button = event.target.closest('[data-action="delete-destination-rule"]');
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const ruleId = button.dataset.ruleId;
  if (!ruleId) {
    return;
  }

  deleteDestinationRule(ruleId);
}

function handleDestinationRulesChange(event) {
  const input = event.target;
  const field = input.dataset.field;
  const ruleId = input.dataset.ruleId;

  if (!field || !ruleId) {
    return;
  }

  const rule = getDestinationRuleById(ruleId);
  if (!rule) {
    return;
  }

  if (field === "label") {
    const nextLabel = normalizeFreeText(input.value);
    if (!nextLabel) {
      render();
      showToast("Le nom du groupe ne peut pas etre vide.", "danger");
      return;
    }
    rule.label = nextLabel;
  }

  if (field === "matchMode") {
    rule.matchMode = normalizeDestinationRuleMatchMode(input.value);
  }

  if (field === "preferredBaqueId") {
    rule.preferredBaqueId = normalizeDestinationRuleTargetBaqueId(input.value);
  }

  if (field === "patterns") {
    const nextPatterns = parseDestinationRulePatterns(input.value);
    if (!nextPatterns.length) {
      render();
      showToast("Ajoutez au moins un mot-cle adresse.", "danger");
      return;
    }
    rule.patterns = nextPatterns;
  }

  rule.updatedAt = new Date().toISOString();
  saveState();
  render();
  showToast("Regle mise a jour.");
}

function handleBaqueGridClick(event) {
  const button = event.target.closest("[data-action]");
  if (!button) {
    return;
  }

  const { action, baqueId, parcelId } = button.dataset;

  if (action === "delete-baque" && baqueId) {
    deleteBaque(baqueId);
  }

  if (action === "toggle-baque-validation" && baqueId) {
    toggleBaqueValidation(baqueId);
  }

  if (action === "delete-parcel" && parcelId) {
    deleteParcel(parcelId);
  }

  if (action === "move-parcel" && parcelId) {
    moveParcel(parcelId);
  }
}

function handleBaqueGridChange(event) {
  const input = event.target;
  const field = input.dataset.field;
  const baqueId = input.dataset.baqueId;

  if (!field || !baqueId) {
    return;
  }

  const baque = getBaqueById(baqueId);
  if (!baque) {
    return;
  }

  const nextValue = input.value.trim();
  if (!nextValue) {
    render();
    showToast("Le champ ne peut pas etre vide.", "danger");
    return;
  }

  baque[field] = nextValue;
  baque.updatedAt = new Date().toISOString();
  saveState();
  render();
  showToast("Baque mise a jour.");
}

function handleModalClick(event) {
  if (event.target instanceof HTMLElement && event.target.dataset.closeWorkspaceCreate === "true") {
    closeWorkspaceCreateModal();
  }

  if (event.target instanceof HTMLElement && event.target.dataset.closeWorkspaceDelete === "true") {
    closeWorkspaceDeleteModal();
  }

  if (event.target instanceof HTMLElement && event.target.dataset.closeScanner === "true") {
    closeScanner();
  }

  if (event.target instanceof HTMLElement && event.target.dataset.closeCapture === "true") {
    closeCaptureModal();
  }
}

function handleCollapseToggle(event) {
  const button = event.target.closest("[data-collapse-toggle]");
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }

  const sectionKey = button.dataset.collapseKey;
  if (!sectionKey || !(sectionKey in collapseState)) {
    return;
  }

  collapseState[sectionKey] = !collapseState[sectionKey];
  saveCollapseState();
  applyCollapseStateToDom();
}

function applyCollapseStateToDom() {
  document.querySelectorAll("[data-collapsible-key]").forEach((section) => {
    if (!(section instanceof HTMLElement)) {
      return;
    }

    const sectionKey = section.dataset.collapsibleKey;
    if (!sectionKey) {
      return;
    }

    const shouldCollapse = Boolean(collapseState[sectionKey]);
    const toggle = section.querySelector("[data-collapse-toggle]");
    const body = section.querySelector(".collapsible-body");

    section.classList.toggle("is-collapsed", shouldCollapse);

    if (toggle instanceof HTMLButtonElement) {
      toggle.textContent = shouldCollapse ? "+" : "-";
      toggle.setAttribute("aria-expanded", String(!shouldCollapse));
      toggle.setAttribute("aria-label", shouldCollapse ? "Ouvrir la rubrique" : "Reduire la rubrique");
      toggle.title = shouldCollapse ? "Ouvrir la rubrique" : "Reduire la rubrique";
    }

    if (body instanceof HTMLElement) {
      body.hidden = shouldCollapse;
    }
  });
}

function openDeliveryNotePicker() {
  ui.deliveryNoteInput.click();
}

async function handleDeliveryNoteImport(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  if (!looksLikePdf(file)) {
    ui.deliveryNoteStatus.textContent = "Ce fichier n'est pas un PDF valide.";
    ui.deliveryNoteInput.value = "";
    showToast("Choisissez un fichier PDF valide.", "danger");
    return;
  }

  const importedAt = new Date().toISOString();
  const deliveryNote = {
    id: createId(),
    name: normalizeFreeText(file.name || "Bon-de-livraison.pdf"),
    size: Number(file.size || 0),
    importedAt,
    updatedAt: importedAt,
    analysis: null,
  };

  try {
    setDeliveryNoteBusy(true, "Import en cours...");
    ui.deliveryNoteStatus.textContent = "Import du PDF en cours...";

    await saveDeliveryNoteFile(deliveryNote.id, file);
    state.deliveryNotes.unshift(deliveryNote);
    saveState();
    renderDeliveryNotes();

    ui.deliveryNoteStatus.textContent = `PDF importe : ${deliveryNote.name}. Analyse en cours...`;

    try {
      await analyzeDeliveryNote(deliveryNote.id, file);
      showToast(`PDF ${deliveryNote.name} importe et compare.`);
    } catch (error) {
      const errorMessage = getDeliveryNoteErrorMessage(error);
      console.error("Erreur d'analyse PDF apres import", error);
      ui.deliveryNoteStatus.textContent = `PDF importe : ${deliveryNote.name}. ${errorMessage}`;
      showToast(errorMessage, "danger");
    }
  } catch (error) {
    const errorMessage = getDeliveryNoteErrorMessage(error, "Impossible d'importer ce PDF.");
    console.error("Erreur d'import PDF", error);
    ui.deliveryNoteStatus.textContent = errorMessage;
    showToast(errorMessage, "danger");
  } finally {
    ui.deliveryNoteInput.value = "";
    setDeliveryNoteBusy(false);
  }
}

async function handleDeliveryNoteListClick(event) {
  const button = event.target.closest("[data-action]");
  if (!(button instanceof HTMLElement)) {
    return;
  }

  if (button.dataset.action === "analyze-delivery-note") {
    const noteId = button.dataset.noteId;
    if (!noteId) {
      return;
    }

    try {
      await analyzeDeliveryNote(noteId);
    } catch (error) {
      const errorMessage = getDeliveryNoteErrorMessage(error);
      console.error("Erreur d'analyse PDF", error);
      ui.deliveryNoteStatus.textContent = errorMessage;
      showToast(errorMessage, "danger");
    }
    return;
  }

  if (button.dataset.action === "simulate-delivery-note") {
    const noteId = button.dataset.noteId;
    if (!noteId) {
      return;
    }

    try {
      await simulateDeliveryNoteParcels(noteId);
    } catch (error) {
      const errorMessage = getDeliveryNoteErrorMessage(error, "Impossible de simuler les colis de ce PDF.");
      console.error("Erreur de simulation PDF", error);
      ui.deliveryNoteStatus.textContent = errorMessage;
      showToast(errorMessage, "danger");
    }
    return;
  }

  if (button.dataset.action !== "delete-delivery-note") {
    return;
  }

  const noteId = button.dataset.noteId;
  const deliveryNote = state.deliveryNotes.find((note) => note.id === noteId);
  if (!noteId || !deliveryNote) {
    return;
  }

  if (!window.confirm(`Supprimer le PDF ${deliveryNote.name} ?`)) {
    return;
  }

  try {
    await deleteDeliveryNoteFile(noteId);
  } catch (error) {
    showToast("Impossible de supprimer ce PDF.", "danger");
    return;
  }

  state.deliveryNotes = state.deliveryNotes.filter((note) => note.id !== noteId);
  saveState();
  renderDeliveryNotes();
  ui.deliveryNoteStatus.textContent = "";
  showToast(`PDF ${deliveryNote.name} supprime.`);
}

async function analyzeDeliveryNote(noteId, providedFile = null) {
  const deliveryNote = state.deliveryNotes.find((note) => note.id === noteId);
  if (!deliveryNote || deliveryNoteAnalysis.busy) {
    return;
  }

  let file = providedFile;
  if (!file) {
    file = await getDeliveryNoteFile(noteId);
  }

  if (!file) {
    throw new Error("missing-pdf-file");
  }

  try {
    setDeliveryNoteBusy(true, "Analyse en cours...");
    deliveryNoteAnalysis.busy = true;
    ui.deliveryNoteStatus.textContent = "Preparation de l'analyse du bon de livraison...";

    const extractedText = await extractTextFromPdfFile(file, (message) => {
      ui.deliveryNoteStatus.textContent = message;
    });
    const entries = parseDeliveryNoteText(extractedText);
    const analyzedAt = new Date().toISOString();
    const analysis = buildStoredDeliveryNoteAnalysis(entries, analyzedAt);

    deliveryNote.analysis = analysis;
    deliveryNote.updatedAt = analyzedAt;

    saveState();
    renderDeliveryNotes();
    ui.deliveryNoteStatus.textContent = analysis.totalMissingCount
      ? `${analysis.totalMissingCount} colis manquants identifies dans ${deliveryNote.name}.`
      : analysis.parseError
      ? `Analyse terminee, mais aucune livraison exploitable n'a ete detectee dans ${deliveryNote.name}.`
      : `Aucun colis manquant detecte dans ${deliveryNote.name}.`;
  } catch (error) {
    const errorMessage = getDeliveryNoteErrorMessage(error);
    const analyzedAt = new Date().toISOString();
    deliveryNote.analysis = {
      totalEntries: 0,
      totalExpectedCount: 0,
      totalRegisteredCount: 0,
      totalMissingCount: 0,
      incomparableParcelsCount: countIncomparableParcels(getPdfComparableParcels()),
      parseError: errorMessage,
      entries: [],
      missingEntries: [],
      analyzedAt,
    };
    deliveryNote.updatedAt = analyzedAt;
    saveState();
    renderDeliveryNotes();
    ui.deliveryNoteStatus.textContent = errorMessage;
    throw error;
  } finally {
    deliveryNoteAnalysis.busy = false;
    setDeliveryNoteBusy(false);
  }
}

async function simulateDeliveryNoteParcels(noteId) {
  const deliveryNote = state.deliveryNotes.find((note) => note.id === noteId);
  if (!deliveryNote || deliveryNoteAnalysis.busy) {
    return;
  }

  const baques = getOrderedBaquesForLayout();
  if (!baques.length) {
    showToast("Ajoutez au moins une baque avant de lancer la simulation.", "danger");
    return;
  }

  const confirmed = window.confirm(
    `Simuler les colis du PDF ${deliveryNote.name} et les repartir aleatoirement dans ${baques.length} baque${baques.length > 1 ? "s" : ""} ? Les colis deja presents sont conserves, mais les memes numeros de commande pourront etre mis a jour et deplaces.`,
  );
  if (!confirmed) {
    return;
  }

  const file = await getDeliveryNoteFile(noteId);
  if (!file) {
    throw new Error("missing-pdf-file");
  }

  try {
    setDeliveryNoteBusy(true, "Simulation...");
    deliveryNoteAnalysis.busy = true;
    ui.deliveryNoteStatus.textContent = `Preparation de la simulation depuis ${deliveryNote.name}...`;

    const extractedText = await extractTextFromPdfFile(file, (message) => {
      ui.deliveryNoteStatus.textContent = message;
    });
    const entries = parseDeliveryNoteText(extractedText);
    if (!entries.length) {
      throw new Error("delivery-note-empty");
    }

    const affectedBaqueIds = new Set();
    let createdCount = 0;
    let updatedCount = 0;

    entries.forEach((entry) => {
      const expectedCount = Math.max(1, Number(entry.expectedCount || 1));

      for (let index = 1; index <= expectedCount; index += 1) {
        const baque = pickRandomBaque(baques);
        const packageIndex = expectedCount > 1 ? `${index}/${expectedCount}` : "";
        const result = upsertSimulatedParcel(
          {
            commandNumber: entry.commandNumber,
            barcode: expectedCount > 1 ? `${entry.commandNumber}${String(index).padStart(3, "0")}` : entry.commandNumber,
            routeCode: "",
            destination: buildDeliveryEntryDestinationLabel(entry),
            client: normalizeFreeText(entry.client || ""),
            description: "",
            routeLabel: "",
            reference: "",
            shippingDate: "",
            weight: "",
            packageIndex,
          },
          baque.id,
        );

        if (result.action === "created") {
          createdCount += 1;
        }

        if (result.action === "updated") {
          updatedCount += 1;
        }

        result.affectedBaqueIds.forEach((baqueId) => {
          if (baqueId) {
            affectedBaqueIds.add(baqueId);
          }
        });
      }
    });

    invalidateBaqueValidations([...affectedBaqueIds]);

    const analyzedAt = new Date().toISOString();
    const analysis = buildStoredDeliveryNoteAnalysis(entries, analyzedAt);
    deliveryNote.analysis = analysis;
    deliveryNote.updatedAt = analyzedAt;

    saveState();
    render();

    const simulatedCount = createdCount + updatedCount;
    if (!simulatedCount) {
      ui.deliveryNoteStatus.textContent = `Aucun colis n'a pu etre simule depuis ${deliveryNote.name}.`;
      showToast("Aucun colis de simulation n'a ete ajoute.", "danger");
      return;
    }

    ui.deliveryNoteStatus.textContent = analysis.totalMissingCount
      ? `Simulation terminee : ${simulatedCount} colis prepares, mais ${analysis.totalMissingCount} colis restent manquants dans ${deliveryNote.name}.`
      : `Simulation terminee : ${simulatedCount} colis prepares depuis ${deliveryNote.name}.`;

    showToast(
      analysis.totalMissingCount
        ? `Simulation terminee : ${createdCount} ajoutes, ${updatedCount} mis a jour. ${analysis.totalMissingCount} colis restent manquants.`
        : `Simulation terminee : ${createdCount} colis ajoutes et ${updatedCount} mis a jour.`,
    );
  } finally {
    deliveryNoteAnalysis.busy = false;
    setDeliveryNoteBusy(false);
  }
}

function openLabelCameraPicker() {
  if (ocr.busy) {
    return;
  }

  void openCaptureModal("label");
}

function openLabelLibraryPicker() {
  if (ocr.busy) {
    return;
  }

  ui.labelLibraryInput.click();
}

function openBarcodeCameraPicker() {
  if (scanner.importingBarcode) {
    return;
  }

  void openCaptureModal("barcode");
}

function openBarcodeLibraryPicker() {
  if (scanner.importingBarcode) {
    return;
  }

  ui.barcodeLibraryInput.click();
}

async function handleBarcodeImageChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  await processBarcodeFile(file);
}

async function processBarcodeFile(file) {
  if (!file) {
    return;
  }

  if (typeof window.Html5Qrcode === "undefined") {
    showToast("La librairie de scan n'a pas pu etre chargee.", "danger");
    resetBarcodeInputs();
    return;
  }

  let fileScanner = null;

  try {
    setBarcodeImportBusy(true);
    if (ui.barcodeStatus) {
      ui.barcodeStatus.textContent = "Analyse de la photo du code-barres...";
    }

    fileScanner = new Html5Qrcode("barcodeFileReader");
    const decodedText = await fileScanner.scanFile(file, false);
    const normalizedCode = decodedText.trim();

    ui.barcodeInput.value = normalizedCode;

    const added = upsertParcel(normalizedCode);
    if (!added) {
      if (ui.barcodeStatus) {
        ui.barcodeStatus.textContent = "Code-barres detecte. Verifiez le numero destination puis enregistrez.";
      }
      showToast("Code-barres detecte.");
    } else {
      if (ui.barcodeStatus) {
        ui.barcodeStatus.textContent = "Code-barres detecte et applique au colis.";
      }
    }
  } catch (error) {
    if (ui.barcodeStatus) {
      ui.barcodeStatus.textContent = "Impossible de lire le code-barres sur cette photo.";
    }
    showToast("Impossible de lire le code-barres. Essayez une photo plus nette.", "danger");
  } finally {
    if (fileScanner) {
      try {
        await fileScanner.clear();
      } catch (error) {
        // Rien a faire si le lecteur fichier est deja nettoye.
      }
    }

    ui.barcodeFileReader.innerHTML = "";
    resetBarcodeInputs();
    setBarcodeImportBusy(false);
  }
}

async function handleLabelImageChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  await processLabelFile(file);
}

async function processLabelFile(file) {
  await processLabelFiles([file]);
}

function scoreParsedLabel(parsed) {
  if (!parsed || typeof parsed !== "object") {
    return Number.NEGATIVE_INFINITY;
  }

  let score = 0;
  if (normalizeRouteCode(parsed.routeCode || "")) {
    score += 140;
  }

  if (getParcelCommandNumber(parsed)) {
    score += 60;
  }

  if (normalizeDestination(parsed.destination || "")) {
    score += 40;
  }

  if (normalizeFreeText(parsed.client || "")) {
    score += 18;
  }

  if (normalizeFreeText(parsed.reference || "")) {
    score += 8;
  }

  if (normalizeFreeText(parsed.weight || "")) {
    score += 6;
  }

  if (normalizeFreeText(parsed.packageIndex || "")) {
    score += 6;
  }

  if (normalizeBarcode(parsed.barcode || "")) {
    score += 3;
  }

  return score;
}

function delay(ms) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

async function processLabelFiles(files) {
  const candidates = (Array.isArray(files) ? files : [files]).filter(Boolean);
  if (!candidates.length) {
    return;
  }

  if (typeof window.Tesseract?.createWorker !== "function") {
    showToast("Le module OCR n'est pas disponible.", "danger");
    resetLabelInputs();
    return;
  }

  try {
    setOcrBusy(true);
    ui.ocrStatus.textContent = "Analyse de l'etiquette en cours...";

    const worker = await getOcrWorker();
    let bestParsed = null;
    let bestScore = Number.NEGATIVE_INFINITY;
    let lastError = null;

    for (let index = 0; index < candidates.length; index += 1) {
      const file = candidates[index];
      ui.ocrStatus.textContent = candidates.length > 1
        ? `Analyse de l'etiquette (${index + 1}/${candidates.length})...`
        : "Analyse de l'etiquette en cours...";

      try {
        const result = await worker.recognize(file);
        const parsed = parseLabelText(result?.data?.text || "");
        const score = scoreParsedLabel(parsed);
        if (score > bestScore) {
          bestParsed = parsed;
          bestScore = score;
        }

        if (parsed.routeCode) {
          break;
        }
      } catch (error) {
        lastError = error;
      }
    }

    if (bestParsed) {
      applyParsedLabelData(bestParsed);
    }

    if (!bestParsed?.routeCode) {
      ui.ocrStatus.textContent = "Lecture terminee, mais le numero destination est introuvable. Essayez une photo plus nette.";
      showToast("Numero destination introuvable. Essayez une photo plus nette.", "danger");
      if (lastError) {
        console.error("OCR etiquette impossible", lastError);
      }
      return;
    }

    ui.ocrStatus.textContent = "Etiquette analysee. Verifiez les champs puis enregistrez le colis.";
    showToast("Etiquette analysee. Les informations ont ete remplies.");
  } catch (error) {
    console.error("OCR etiquette impossible", error);
    ui.ocrStatus.textContent = "Impossible de lire l'etiquette. Essayez une photo bien droite et nette.";
    showToast("Impossible de lire l'etiquette. Reessayez avec une photo plus nette.", "danger");
  } finally {
    setOcrBusy(false);
    resetLabelInputs();
  }
}

async function openCaptureModal(mode) {
  if (captureSession.busy) {
    return;
  }

  if (!navigator.mediaDevices?.getUserMedia) {
    fallbackToNativeCapture(mode);
    return;
  }

  if (scanner.active) {
    await closeScanner();
  }

  captureSession.mode = mode;
  configureCaptureModal(mode);
  ui.captureStatus.textContent = "Demande d'acces a la camera...";
  ui.takeCaptureBtn.disabled = true;
  ui.captureModal.classList.remove("hidden");
  ui.captureModal.setAttribute("aria-hidden", "false");

  try {
    await startCaptureStream();
    ui.captureStatus.textContent = mode === "label"
      ? "Cadrez l'etiquette tranquillement. La capture se lance apres un court maintien."
      : "Centrez le code-barres dans le rectangle, puis prenez la photo.";
  } catch (error) {
    await closeCaptureModal({ silent: true });
    showToast("Impossible d'ouvrir la camera integree. Utilisez Choisir une photo si besoin.", "danger");
  }
}

function configureCaptureModal(mode) {
  const isLabelMode = mode === "label";

  ui.captureTitle.textContent = isLabelMode ? "Cadrer l'etiquette" : "Cadrer le code-barres";
  ui.captureHint.textContent = isLabelMode
    ? "Prenez votre temps pour cadrer. Quand l'etiquette reste stable, l'app prend plusieurs photos."
    : "Placez le code-barres au centre du cadre, evitez les reflets, puis prenez la photo.";
  ui.takeCaptureBtn.textContent = isLabelMode ? "Prendre maintenant" : "Prendre la photo";
  ui.captureGuide.classList.toggle("capture-guide--label", isLabelMode);
  ui.captureGuide.classList.toggle("capture-guide--barcode", !isLabelMode);
  setCaptureDetectionState(false);
}

async function startCaptureStream() {
  await stopCaptureStream();

  const stream = await navigator.mediaDevices.getUserMedia({
    video: {
      facingMode: { ideal: "environment" },
      width: { ideal: 1920 },
      height: { ideal: 1080 },
    },
    audio: false,
  });

  captureSession.stream = stream;
  ui.captureVideo.srcObject = stream;

  await ui.captureVideo.play();
  ui.takeCaptureBtn.disabled = false;
  if (captureSession.mode === "label") {
    startAutoCaptureMonitoring();
  }
}

async function closeCaptureModal(options = {}) {
  if (captureSession.busy && !options.force) {
    return;
  }

  await stopCaptureStream();
  ui.captureModal.classList.add("hidden");
  ui.captureModal.setAttribute("aria-hidden", "true");

  if (!options.silent) {
    ui.captureStatus.textContent = "";
  }
}

async function stopCaptureStream() {
  stopAutoCaptureMonitoring();

  if (captureSession.stream) {
    captureSession.stream.getTracks().forEach((track) => track.stop());
    captureSession.stream = null;
  }

  if (ui.captureVideo) {
    ui.captureVideo.pause();
    ui.captureVideo.srcObject = null;
  }
}

async function handleCapturePhoto(options = {}) {
  if (captureSession.busy || !ui.captureVideo.videoWidth || !ui.captureVideo.videoHeight) {
    return;
  }

  captureSession.busy = true;
  stopAutoCaptureMonitoring();
  ui.takeCaptureBtn.disabled = true;
  ui.takeCaptureBtn.textContent = options.auto ? "Capture auto..." : "Preparation...";
  ui.captureStatus.textContent = options.auto
    ? "Etiquette stable detectee. Preparation de l'image..."
    : "Photo prise. Preparation de l'image...";

  const mode = captureSession.mode;

  try {
    const files = mode === "label"
      ? await captureLabelBurst()
      : [await captureCurrentFrame(mode)];
    await closeCaptureModal({ force: true, silent: true });

    if (mode === "label") {
      await processLabelFiles(files);
    } else {
      await processBarcodeFile(files[0]);
    }
  } catch (error) {
    ui.captureStatus.textContent = "Impossible de prendre la photo. Reessayez avec une image plus stable.";
    showToast("Impossible de prendre la photo. Reessayez.", "danger");
  } finally {
    captureSession.busy = false;
    ui.takeCaptureBtn.disabled = false;
    ui.takeCaptureBtn.textContent = mode === "label" ? "Prendre maintenant" : "Prendre la photo";
  }
}

function startAutoCaptureMonitoring() {
  stopAutoCaptureMonitoring();

  captureSession.autoTriggered = false;
  captureSession.stableFrameCount = 0;
  captureSession.lastFrameSignature = null;
  ui.captureStatus.textContent = "Auto actif. Cadrez l'etiquette et gardez-la stable un instant.";

  captureSession.autoCaptureTimer = window.setInterval(() => {
    if (captureSession.mode !== "label" || captureSession.busy || captureSession.autoTriggered) {
      return;
    }

    if (!ui.captureVideo.videoWidth || !ui.captureVideo.videoHeight || ui.captureVideo.readyState < 2) {
      return;
    }

    const analysis = analyzeCurrentCaptureFrame();
    if (!analysis) {
      return;
    }

    captureSession.lastFrameSignature = analysis.signature;

    if (!analysis.isReady) {
      captureSession.stableFrameCount = 0;
      setCaptureDetectionState(false);
      ui.captureStatus.textContent = analysis.message;
      return;
    }

    setCaptureDetectionState(true);

    if (analysis.motion !== null && analysis.motion > 14) {
      captureSession.stableFrameCount = 0;
      ui.captureStatus.textContent = analysis.message;
      return;
    }

    captureSession.stableFrameCount += 1;

    if (captureSession.stableFrameCount < LABEL_AUTO_CAPTURE_STABLE_FRAMES) {
      const remainingFrames = LABEL_AUTO_CAPTURE_STABLE_FRAMES - captureSession.stableFrameCount;
      const remainingSeconds = Math.max(1, Math.ceil((remainingFrames * LABEL_AUTO_CAPTURE_POLL_MS) / 1000));
      ui.captureStatus.textContent = remainingSeconds > 1
        ? `Etiquette detectee. Gardez le telephone stable encore ${remainingSeconds} secondes.`
        : "Etiquette detectee. Gardez le telephone stable encore un instant.";
      return;
    }

    captureSession.autoTriggered = true;
    ui.captureStatus.textContent = "Etiquette stable. Rafale automatique en preparation...";
    captureSession.autoCaptureKickoffTimer = window.setTimeout(() => {
      captureSession.autoCaptureKickoffTimer = 0;
      if (captureSession.mode === "label" && !captureSession.busy) {
        void handleCapturePhoto({ auto: true });
      }
    }, LABEL_AUTO_CAPTURE_KICKOFF_MS);
  }, LABEL_AUTO_CAPTURE_POLL_MS);
}

function stopAutoCaptureMonitoring() {
  if (captureSession.autoCaptureTimer) {
    window.clearInterval(captureSession.autoCaptureTimer);
    captureSession.autoCaptureTimer = 0;
  }

  if (captureSession.autoCaptureKickoffTimer) {
    window.clearTimeout(captureSession.autoCaptureKickoffTimer);
    captureSession.autoCaptureKickoffTimer = 0;
  }

  captureSession.lastFrameSignature = null;
  captureSession.stableFrameCount = 0;
  captureSession.autoTriggered = false;
  setCaptureDetectionState(false);
}

function analyzeCurrentCaptureFrame() {
  const crop = getCaptureCropArea("label");
  if (!crop.width || !crop.height) {
    return null;
  }

  const sampleWidth = 108;
  const sampleHeight = Math.max(68, Math.round(sampleWidth * (crop.height / crop.width)));
  const { context } = getCaptureAnalysisContext(sampleWidth, sampleHeight);
  context.drawImage(
    ui.captureVideo,
    crop.x,
    crop.y,
    crop.width,
    crop.height,
    0,
    0,
    sampleWidth,
    sampleHeight,
  );

  const imageData = context.getImageData(0, 0, sampleWidth, sampleHeight);
  const signature = new Uint8Array(sampleWidth * sampleHeight);
  const totalPixels = sampleWidth * sampleHeight;
  let whitePixels = 0;
  let centerWhitePixels = 0;
  let centerPixelCount = 0;
  let luminanceSum = 0;
  let luminanceSquaredSum = 0;
  let edgePixels = 0;

  for (let y = 0; y < sampleHeight; y += 1) {
    for (let x = 0; x < sampleWidth; x += 1) {
      const pixelIndex = (y * sampleWidth + x);
      const offset = pixelIndex * 4;
      const red = imageData.data[offset];
      const green = imageData.data[offset + 1];
      const blue = imageData.data[offset + 2];
      const luminance = Math.round((red * 0.299) + (green * 0.587) + (blue * 0.114));
      const colorSpread = Math.max(red, green, blue) - Math.min(red, green, blue);
      const isWhiteLabelPixel = luminance >= 168 && colorSpread <= 78;

      signature[pixelIndex] = luminance;
      luminanceSum += luminance;
      luminanceSquaredSum += luminance * luminance;

      if (isWhiteLabelPixel) {
        whitePixels += 1;
      }

      if (
        x >= Math.round(sampleWidth * 0.2)
        && x <= Math.round(sampleWidth * 0.8)
        && y >= Math.round(sampleHeight * 0.2)
        && y <= Math.round(sampleHeight * 0.8)
      ) {
        centerPixelCount += 1;
        if (isWhiteLabelPixel) {
          centerWhitePixels += 1;
        }
      }

      if (x > 0 && y > 0) {
        const leftLuminance = signature[pixelIndex - 1];
        const topLuminance = signature[pixelIndex - sampleWidth];
        if (Math.abs(luminance - leftLuminance) + Math.abs(luminance - topLuminance) >= 54) {
          edgePixels += 1;
        }
      }
    }
  }

  const averageLuminance = luminanceSum / totalPixels;
  const variance = Math.max(0, (luminanceSquaredSum / totalPixels) - (averageLuminance * averageLuminance));
  const contrast = Math.sqrt(variance);
  const whiteRatio = whitePixels / totalPixels;
  const centerWhiteRatio = centerPixelCount ? centerWhitePixels / centerPixelCount : 0;
  const edgeRatio = edgePixels / totalPixels;
  const motion = getCaptureFrameMotion(signature, captureSession.lastFrameSignature);
  const isReady = whiteRatio >= 0.34 && centerWhiteRatio >= 0.52 && contrast >= 34 && edgeRatio >= 0.06;

  return {
    signature,
    isReady,
    message: buildAutoCaptureStatusMessage({
      whiteRatio,
      centerWhiteRatio,
      contrast,
      edgeRatio,
      motion,
    }),
  };
}

function setCaptureDetectionState(isDetected) {
  ui.captureGuide?.classList.toggle("capture-guide--detected", isDetected);
  ui.capturePreview?.classList.toggle("capture-preview--detected", isDetected);
}

function getCaptureAnalysisContext(width, height) {
  if (!captureSession.analysisCanvas) {
    captureSession.analysisCanvas = document.createElement("canvas");
    captureSession.analysisContext = captureSession.analysisCanvas.getContext("2d", { willReadFrequently: true });
  }

  if (!captureSession.analysisContext) {
    throw new Error("capture-analysis-context-unavailable");
  }

  if (captureSession.analysisCanvas.width !== width || captureSession.analysisCanvas.height !== height) {
    captureSession.analysisCanvas.width = width;
    captureSession.analysisCanvas.height = height;
  }

  return {
    canvas: captureSession.analysisCanvas,
    context: captureSession.analysisContext,
  };
}

function getCaptureFrameMotion(currentSignature, previousSignature) {
  if (!(previousSignature instanceof Uint8Array) || previousSignature.length !== currentSignature.length) {
    return null;
  }

  let totalDelta = 0;
  let comparedPixels = 0;
  for (let index = 0; index < currentSignature.length; index += 3) {
    totalDelta += Math.abs(currentSignature[index] - previousSignature[index]);
    comparedPixels += 1;
  }

  return comparedPixels ? totalDelta / comparedPixels : null;
}

function buildAutoCaptureStatusMessage(metrics) {
  if (metrics.centerWhiteRatio < 0.52 || metrics.whiteRatio < 0.34) {
    return "Auto actif. Rapprochez ou recentrez l'etiquette pour qu'elle remplisse mieux le cadre.";
  }

  if (metrics.contrast < 34 || metrics.edgeRatio < 0.06) {
    return "Auto actif. Evitez le flou et les reflets sur l'etiquette.";
  }

  if (metrics.motion !== null && metrics.motion > 16) {
    return "Etiquette detectee. Restez bien dans l'axe sans bouger, la rafale arrive.";
  }

  return "Etiquette detectee. Gardez la position encore un instant.";
}

async function captureLabelBurst() {
  const captured = [];

  for (let index = 0; index < LABEL_BURST_COUNT; index += 1) {
    ui.captureStatus.textContent = `Rafale en cours (${index + 1}/${LABEL_BURST_COUNT})... gardez l'etiquette dans le cadre.`;
    if (index) {
      await delay(LABEL_BURST_INTERVAL_MS);
    }
    captured.push(await captureCurrentFrame("label"));
  }

  return captured;
}

async function captureCurrentFrame(mode) {
  const crop = getCaptureCropArea(mode);
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("canvas-unavailable");
  }

  canvas.width = Math.max(1, Math.round(crop.width));
  canvas.height = Math.max(1, Math.round(crop.height));
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = "high";
  if (mode === "label") {
    context.filter = "grayscale(1) contrast(1.35) brightness(1.06)";
  }
  context.drawImage(
    ui.captureVideo,
    crop.x,
    crop.y,
    crop.width,
    crop.height,
    0,
    0,
    canvas.width,
    canvas.height,
  );
  context.filter = "none";

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob(
      (result) => {
        if (result) {
          resolve(result);
        } else {
          reject(new Error("blob-unavailable"));
        }
      },
      "image/jpeg",
      0.92,
    );
  });

  return new File([blob], `${mode}-${Date.now()}.jpg`, { type: "image/jpeg" });
}

function getCaptureCropArea(mode) {
  const sourceWidth = ui.captureVideo.videoWidth;
  const sourceHeight = ui.captureVideo.videoHeight;
  const videoRect = ui.captureVideo.getBoundingClientRect();
  const guideRect = ui.captureGuide.getBoundingClientRect();
  const displayWidth = videoRect.width;
  const displayHeight = videoRect.height;

  if (!sourceWidth || !sourceHeight || !displayWidth || !displayHeight) {
    return { x: 0, y: 0, width: sourceWidth, height: sourceHeight };
  }

  const scale = Math.max(displayWidth / sourceWidth, displayHeight / sourceHeight);
  const renderedWidth = sourceWidth * scale;
  const renderedHeight = sourceHeight * scale;
  const overflowX = Math.max(0, (renderedWidth - displayWidth) / 2);
  const overflowY = Math.max(0, (renderedHeight - displayHeight) / 2);

  let x = (guideRect.left - videoRect.left + overflowX) / scale;
  let y = (guideRect.top - videoRect.top + overflowY) / scale;
  let width = guideRect.width / scale;
  let height = guideRect.height / scale;

  const extraMargin = mode === "label" ? 0.12 : 0.1;
  x -= width * extraMargin;
  y -= height * extraMargin;
  width *= 1 + extraMargin * 2;
  height *= 1 + extraMargin * 2;

  x = clamp(x, 0, sourceWidth - 1);
  y = clamp(y, 0, sourceHeight - 1);
  width = clamp(width, 1, sourceWidth - x);
  height = clamp(height, 1, sourceHeight - y);

  return { x, y, width, height };
}

function fallbackToNativeCapture(mode) {
  if (mode === "label") {
    ui.labelCameraInput.click();
  } else {
    ui.barcodeCameraInput.click();
  }
}

function setDeliveryNoteBusy(isBusy, busyLabel = "Import en cours...") {
  ui.importDeliveryNoteBtn.disabled = isBusy;
  ui.importDeliveryNoteBtn.textContent = isBusy ? busyLabel : "Importer un PDF";
}

async function extractTextFromPdfFile(file, onProgress) {
  const pdfjsLib = await getPdfJs();
  const data = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdfDocument = await loadingTask.promise;

  try {
    const embeddedText = await extractEmbeddedTextFromPdfDocument(pdfDocument, onProgress);
    if (parseDeliveryNoteText(embeddedText).length) {
      return embeddedText;
    }

    if (typeof window.Tesseract?.createWorker !== "function") {
      throw new Error("tesseract-unavailable");
    }

    return await extractOcrTextFromPdfDocument(pdfDocument, onProgress);
  } finally {
    await loadingTask.destroy();
  }
}

async function extractEmbeddedTextFromPdfDocument(pdfDocument, onProgress) {
  const textChunks = [];

  for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
    onProgress?.(`Lecture du texte du PDF page ${pageNumber}/${pdfDocument.numPages}...`);

    const page = await pdfDocument.getPage(pageNumber);

    try {
      const lines = await extractPdfPageLines(page);
      textChunks.push(lines.join("\n"));
    } finally {
      page.cleanup();
    }
  }

  return textChunks.join("\n");
}

async function extractOcrTextFromPdfDocument(pdfDocument, onProgress) {
  const textChunks = [];

  for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
    onProgress?.(`Analyse OCR du PDF page ${pageNumber}/${pdfDocument.numPages}...`);

    const page = await pdfDocument.getPage(pageNumber);

    try {
      const pageImage = await renderPdfPageToCanvas(page);
      const worker = await getOcrWorker();
      const result = await worker.recognize(pageImage);
      textChunks.push(result.data.text || "");
    } finally {
      page.cleanup();
    }
  }

  return textChunks.join("\n");
}

async function extractPdfPageLines(page) {
  const textContent = await page.getTextContent();
  return groupPdfTextItemsIntoLines(textContent.items || []);
}

function groupPdfTextItemsIntoLines(items) {
  const positionedItems = items
    .filter((item) => item?.str && item.str.trim())
    .map((item) => ({
      text: item.str,
      x: Number(item.transform?.[4] || 0),
      y: Number(item.transform?.[5] || 0),
      width: Number(item.width || 0),
    }));
  const rows = [];
  const lineTolerance = 1.5;

  positionedItems.forEach((item) => {
    let row = rows.find((candidate) => Math.abs(candidate.y - item.y) <= lineTolerance);
    if (!row) {
      row = { y: item.y, items: [] };
      rows.push(row);
    }

    row.items.push(item);
  });

  return rows
    .sort((left, right) => right.y - left.y)
    .map((row) => {
      row.items.sort((left, right) => left.x - right.x);

      let line = "";
      let previousEnd = null;

      row.items.forEach((item) => {
        const gap = previousEnd === null ? 0 : item.x - previousEnd;
        if (line && gap > 1.5) {
          line += " ";
        }

        line += item.text;
        previousEnd = item.x + item.width;
      });

      return normalizeDeliveryTextLine(line);
    })
    .filter(Boolean);
}

function setOcrBusy(isBusy) {
  ocr.busy = isBusy;
  ui.scanLabelBtn.disabled = isBusy;
  ui.chooseLabelBtn.disabled = isBusy;
  ui.scanLabelBtn.textContent = isBusy ? "Analyse en cours..." : "Prendre une photo";
  ui.chooseLabelBtn.textContent = isBusy ? "Analyse en cours..." : "Choisir une photo";
}

function setBarcodeImportBusy(isBusy) {
  scanner.importingBarcode = isBusy;
  if (ui.importBarcodeBtn) {
    ui.importBarcodeBtn.disabled = isBusy;
    ui.importBarcodeBtn.textContent = isBusy ? "Analyse en cours..." : "Prendre une photo";
  }
  if (ui.chooseBarcodeBtn) {
    ui.chooseBarcodeBtn.disabled = isBusy;
    ui.chooseBarcodeBtn.textContent = isBusy ? "Analyse en cours..." : "Choisir une photo";
  }
}

function resetLabelInputs() {
  ui.labelCameraInput.value = "";
  ui.labelLibraryInput.value = "";
}

function resetBarcodeInputs() {
  ui.barcodeCameraInput.value = "";
  ui.barcodeLibraryInput.value = "";
}

async function getOcrWorker() {
  if (ocr.worker) {
    return ocr.worker;
  }

  ocr.worker = await Tesseract.createWorker("fra+eng");
  return ocr.worker;
}

async function stopOcrWorker() {
  if (!ocr.worker) {
    return;
  }

  try {
    await ocr.worker.terminate();
  } catch (error) {
    // Le worker peut deja etre arrete.
  }

  ocr.worker = null;
}

function applyParsedLabelData(parsed) {
  if (parsed.routeCode) {
    ui.routeCodeInput.value = parsed.routeCode;
  }

  if (parsed.destination) {
    ui.destinationInput.value = parsed.destination;
  }

  if (parsed.client) {
    ui.clientInput.value = parsed.client;
  }

  if (parsed.description) {
    ui.descriptionInput.value = parsed.description;
  }

  if (parsed.routeLabel) {
    ui.routeLabelInput.value = parsed.routeLabel;
  }

  if (parsed.reference) {
    ui.referenceInput.value = parsed.reference;
  }

  if (parsed.shippingDate) {
    ui.shippingDateInput.value = parsed.shippingDate;
  }

  if (parsed.weight) {
    ui.weightInput.value = parsed.weight;
  }

  if (parsed.measuredDimensions) {
    ui.measuredDimensionsInput.value = parsed.measuredDimensions;
  }

  if (parsed.packageIndex) {
    ui.packageIndexInput.value = parsed.packageIndex;
  }

  if (parsed.barcode) {
    ui.barcodeInput.value = parsed.barcode;
  }
}

function upsertParcel(scannedBarcode = "") {
  const baqueId = ui.parcelBaqueSelect.value;
  const routeCode = normalizeRouteCode(ui.routeCodeInput.value);
  const destination = normalizeDestination(ui.destinationInput.value);
  const client = normalizeFreeText(ui.clientInput.value);
  const description = normalizeFreeText(ui.descriptionInput.value);
  const routeLabel = normalizeFreeText(ui.routeLabelInput.value);
  const reference = normalizeFreeText(ui.referenceInput.value);
  const shippingDate = normalizeFreeText(ui.shippingDateInput.value);
  const weight = normalizeFreeText(ui.weightInput.value);
  const measuredDimensions = normalizeFreeText(ui.measuredDimensionsInput.value);
  const packageIndex = normalizeFreeText(ui.packageIndexInput.value);
  const barcode = normalizeBarcode(scannedBarcode || ui.barcodeInput.value);
  const normalizedParcelData = normalizeParcelData({
    barcode,
    routeCode,
    destination,
    client,
    description,
    routeLabel,
    reference,
    shippingDate,
    weight,
    measuredDimensions,
    packageIndex,
  });

  if (!baqueId || !normalizedParcelData.routeCode) {
    showToast("Choisissez une baque et renseignez le numero destination.", "danger");
    return false;
  }

  const baque = getBaqueById(baqueId);
  if (!baque) {
    showToast("La baque selectionnee est introuvable.", "danger");
    return false;
  }

  const now = new Date().toISOString();
  const existing = findExistingParcel(state.parcels, normalizedParcelData);
  const commandNumber = getParcelCommandNumber(normalizedParcelData);

  if (existing) {
    const moved = existing.currentBaqueId !== baqueId;
    const previousBaqueId = existing.currentBaqueId;
    existing.barcode = normalizedParcelData.barcode;
    existing.commandNumber = commandNumber;
    existing.routeCode = normalizedParcelData.routeCode;
    existing.destination = normalizedParcelData.destination;
    existing.client = normalizedParcelData.client;
    existing.description = normalizedParcelData.description;
    existing.routeLabel = normalizedParcelData.routeLabel;
    existing.reference = normalizedParcelData.reference;
    existing.shippingDate = normalizedParcelData.shippingDate;
    existing.weight = normalizedParcelData.weight;
    existing.measuredDimensions = normalizedParcelData.measuredDimensions;
    existing.packageIndex = normalizedParcelData.packageIndex;
    existing.currentBaqueId = baqueId;
    existing.updatedAt = now;

    invalidateBaqueValidations([previousBaqueId, baqueId]);
    saveState();
    render();
    clearParcelForm();
    showToast(
      moved
        ? `Colis ${getParcelIdentifier(normalizedParcelData)} deplace vers ${baque.name}.`
        : `Colis ${getParcelIdentifier(normalizedParcelData)} mis a jour.`,
    );
    return true;
  }

  state.parcels.unshift({
    id: createId(),
    barcode: normalizedParcelData.barcode,
    commandNumber,
    routeCode: normalizedParcelData.routeCode,
    destination: normalizedParcelData.destination,
    client: normalizedParcelData.client,
    description: normalizedParcelData.description,
    routeLabel: normalizedParcelData.routeLabel,
    reference: normalizedParcelData.reference,
    shippingDate: normalizedParcelData.shippingDate,
    weight: normalizedParcelData.weight,
    measuredDimensions: normalizedParcelData.measuredDimensions,
    packageIndex: normalizedParcelData.packageIndex,
    currentBaqueId: baqueId,
    originBaqueId: baqueId,
    originBaqueLabel: baque.name,
    createdAt: now,
    updatedAt: now,
  });

  invalidateBaqueValidation(baqueId);
  saveState();
  render();
  clearParcelForm();
  showToast(
    commandNumber
      ? `Colis ${getParcelIdentifier({ ...normalizedParcelData, commandNumber })} ajoute dans ${baque.name}.`
      : `Colis ${getParcelIdentifier(normalizedParcelData)} ajoute dans ${baque.name}. Il ne pourra pas etre compare au PDF sans numero de commande.`,
  );
  return true;
}

function clearParcelForm() {
  ui.routeCodeInput.value = "";
  ui.destinationInput.value = "";
  ui.clientInput.value = "";
  ui.descriptionInput.value = "";
  ui.routeLabelInput.value = "";
  ui.referenceInput.value = "";
  ui.shippingDateInput.value = "";
  ui.weightInput.value = "";
  ui.measuredDimensionsInput.value = "";
  ui.packageIndexInput.value = "";
  ui.barcodeInput.value = "";
  ui.ocrStatus.textContent = "";
  if (ui.barcodeStatus) {
    ui.barcodeStatus.textContent = "";
  }
  ui.routeCodeInput.focus();
}

function upsertSmallParcelScan(scannedCode = "") {
  const normalizedCode = normalizeBarcode(scannedCode || ui.smallParcelBarcodeInput?.value || "");
  const quantity = clamp(Math.round(Number(ui.smallParcelQuantityInput?.value || 1)), 1, 99);
  const normalizedScan = normalizeSmallParcelScan({
    barcode: normalizedCode,
    commandNumber: normalizedCode,
    quantity,
  });

  if (!normalizedScan?.commandNumber) {
    showToast("Scannez un numero de commande ou un code-barres exploitable.", "danger");
    return false;
  }

  const now = new Date().toISOString();
  state.smallParcelScans.unshift({
    ...normalizedScan,
    id: createId(),
    createdAt: now,
    updatedAt: now,
  });

  saveState();
  renderHeroStats();
  renderSmallParcelScans();
  renderDeliveryNotes();
  clearSmallParcelForm();
  showToast(
    `${quantity} ${pluralize(quantity, "petit colis compte", "petits colis comptes")} pour la commande ${normalizedScan.commandNumber}.`,
  );
  return true;
}

function clearSmallParcelForm() {
  if (ui.smallParcelBarcodeInput) {
    ui.smallParcelBarcodeInput.value = "";
  }
  if (ui.smallParcelQuantityInput) {
    ui.smallParcelQuantityInput.value = "1";
  }
  ui.smallParcelBarcodeInput?.focus();
}

function deleteSmallParcelScan(scanId) {
  const scan = state.smallParcelScans.find((item) => item.id === scanId);
  if (!scan) {
    return;
  }

  const quantity = Math.max(1, Number(scan.quantity || 1));
  if (!window.confirm(`Supprimer ${quantity} ${pluralize(quantity, "petit colis", "petits colis")} de la commande ${scan.commandNumber} ?`)) {
    return;
  }

  state.smallParcelScans = state.smallParcelScans.filter((item) => item.id !== scanId);
  saveState();
  renderHeroStats();
  renderSmallParcelScans();
  renderDeliveryNotes();
  showToast(`Scan petits colis ${scan.commandNumber} supprime.`);
}

function deleteBaque(baqueId) {
  if (state.baques.length === 1) {
    showToast("Vous devez garder au moins une baque.", "danger");
    return;
  }

  const baque = getBaqueById(baqueId);
  const parcelCount = state.parcels.filter((parcel) => parcel.currentBaqueId === baqueId).length;
  const confirmed = window.confirm(
    `Supprimer ${baque?.name || "cette baque"} et ${parcelCount} colis qu'elle contient ?`,
  );

  if (!confirmed) {
    return;
  }

  const now = new Date().toISOString();
  state.baques = state.baques.filter((item) => item.id !== baqueId);
  state.parcels = state.parcels.filter((parcel) => parcel.currentBaqueId !== baqueId);
  state.destinationRules = state.destinationRules.map((rule) => (
    rule.preferredBaqueId === baqueId
      ? { ...rule, preferredBaqueId: "", updatedAt: now }
      : rule
  ));
  saveState();
  render();
  showToast("Baque supprimee.");
}

function deleteDestinationRule(ruleId) {
  const rule = getDestinationRuleById(ruleId);
  if (!rule) {
    return;
  }

  if (!window.confirm(`Supprimer la regle "${rule.label}" ?`)) {
    return;
  }

  state.destinationRules = state.destinationRules.filter((item) => item.id !== ruleId);
  saveState();
  render();
  showToast(`Regle "${rule.label}" supprimee.`);
}

function toggleBaqueValidation(baqueId) {
  const baque = getBaqueById(baqueId);
  if (!baque) {
    return;
  }

  if (baque.validatedAt) {
    baque.validatedAt = "";
    baque.updatedAt = new Date().toISOString();
    saveState();
    render();
    showToast(`Validation retiree pour ${baque.name}.`);
    return;
  }

  baque.validatedAt = new Date().toISOString();
  baque.updatedAt = baque.validatedAt;
  saveState();
  render();
  showToast(`${baque.name} marquee comme terminee.`);
}

function invalidateBaqueValidation(baqueId) {
  if (!baqueId) {
    return false;
  }

  const baque = getBaqueById(baqueId);
  if (!baque?.validatedAt) {
    return false;
  }

  baque.validatedAt = "";
  baque.updatedAt = new Date().toISOString();
  return true;
}

function invalidateBaqueValidations(baqueIds) {
  const uniqueBaqueIds = [...new Set(
    (Array.isArray(baqueIds) ? baqueIds : [baqueIds])
      .filter(Boolean),
  )];
  let invalidated = false;

  uniqueBaqueIds.forEach((baqueId) => {
    invalidated = invalidateBaqueValidation(baqueId) || invalidated;
  });

  return invalidated;
}

function deleteParcel(parcelId) {
  const parcel = state.parcels.find((item) => item.id === parcelId);
  if (!parcel) {
    return;
  }

  if (!window.confirm(`Supprimer le colis ${getParcelIdentifier(parcel)} ?`)) {
    return;
  }

  invalidateBaqueValidation(parcel.currentBaqueId);
  state.parcels = state.parcels.filter((item) => item.id !== parcelId);
  saveState();
  render();
  showToast(`Colis ${getParcelIdentifier(parcel)} supprime.`);
}

function moveParcel(parcelId) {
  const parcel = state.parcels.find((item) => item.id === parcelId);
  if (!parcel) {
    return;
  }

  const select = document.querySelector(
    `[data-role="move-select"][data-parcel-id="${CSS.escape(parcelId)}"]`,
  );

  if (!(select instanceof HTMLSelectElement)) {
    return;
  }

  const nextBaqueId = select.value;
  const nextBaque = getBaqueById(nextBaqueId);

  if (!nextBaque) {
    showToast("La baque cible est introuvable.", "danger");
    return;
  }

  if (parcel.currentBaqueId === nextBaqueId) {
    showToast("Ce colis est deja dans cette baque.");
    return;
  }

  const previousBaqueId = parcel.currentBaqueId;
  parcel.currentBaqueId = nextBaqueId;
  parcel.updatedAt = new Date().toISOString();
  invalidateBaqueValidations([previousBaqueId, nextBaqueId]);
  saveState();
  render();
  showToast(`Colis ${getParcelIdentifier(parcel)} deplace vers ${nextBaque.name}.`);
}

async function openScanner(target = "parcel") {
  if (typeof window.Html5QrcodeScanner === "undefined") {
    showToast("La librairie de scan n'a pas pu etre chargee.", "danger");
    return;
  }

  scanner.target = target === "small-parcel" ? "small-parcel" : "parcel";
  if (scanner.active) {
    ui.scannerModal.classList.remove("hidden");
    ui.scannerModal.setAttribute("aria-hidden", "false");
    return;
  }

  scanner.handled = false;
  if (ui.scannerHelp) {
    ui.scannerHelp.textContent = scanner.target === "small-parcel"
      ? "Le code scanne alimentera la section Petits colis a part et servira uniquement au comptage du bon de livraison."
      : "Le colis sera ajoute automatiquement si la baque et le numero destination sont deja remplis. Si besoin, vous pouvez aussi choisir une image du code-barres.";
  }
  ui.scannerStatus.textContent = "Autorisez la camera ou utilisez le scan par image si besoin.";
  ui.scannerModal.classList.remove("hidden");
  ui.scannerModal.setAttribute("aria-hidden", "false");

  try {
    scanner.instance = new Html5QrcodeScanner(
      "reader",
      {
        fps: 10,
        qrbox: { width: 320, height: 140 },
        aspectRatio: 1.7777778,
        rememberLastUsedCamera: true,
        showTorchButtonIfSupported: true,
        supportedScanTypes: [
          Html5QrcodeScanType.SCAN_TYPE_CAMERA,
          Html5QrcodeScanType.SCAN_TYPE_FILE,
        ],
        formatsToSupport: [
          Html5QrcodeSupportedFormats.CODE_128,
          Html5QrcodeSupportedFormats.CODE_39,
          Html5QrcodeSupportedFormats.CODE_93,
          Html5QrcodeSupportedFormats.EAN_13,
          Html5QrcodeSupportedFormats.EAN_8,
          Html5QrcodeSupportedFormats.UPC_A,
          Html5QrcodeSupportedFormats.UPC_E,
          Html5QrcodeSupportedFormats.ITF,
          Html5QrcodeSupportedFormats.CODABAR,
        ],
      },
      false,
    );

    scanner.instance.render(
      async (decodedText) => {
        if (scanner.handled) {
          return;
        }

        scanner.handled = true;
        const normalizedText = decodedText.trim();
        const added = scanner.target === "small-parcel"
          ? (() => {
            if (ui.smallParcelBarcodeInput) {
              ui.smallParcelBarcodeInput.value = normalizedText;
            }
            return upsertSmallParcelScan(normalizedText);
          })()
          : (() => {
            ui.barcodeInput.value = normalizedText;
            return upsertParcel(normalizedText);
          })();
        if (!added) {
          showToast(
            scanner.target === "small-parcel"
              ? "Code detecte. Completez la quantite puis validez."
              : "Code detecte. Completez les champs puis validez.",
          );
        }

        await closeScanner();
      },
      () => {},
    );

    scanner.active = true;
    ui.scannerStatus.textContent = scanner.target === "small-parcel"
      ? "Scanner actif pour les petits colis. Visez le code de commande."
      : "Scanner actif. Vous pouvez utiliser la camera ou importer une photo du code.";
  } catch (error) {
    await stopScanner();
    ui.scannerModal.classList.add("hidden");
    ui.scannerModal.setAttribute("aria-hidden", "true");
    showToast("Impossible de lancer le scanner. Verifiez les permissions camera.", "danger");
  }
}

async function closeScanner() {
  await stopScanner();
  ui.scannerModal.classList.add("hidden");
  ui.scannerModal.setAttribute("aria-hidden", "true");
  ui.scannerStatus.textContent = "";
}

async function stopScanner() {
  scanner.active = false;
  scanner.handled = false;

  if (scanner.instance) {
    try {
      await scanner.instance.clear();
    } catch (error) {
      // Le scanner peut deja etre libere si l'utilisateur a ferme la camera.
    }

    scanner.instance = null;
  }

  if (ui.reader) {
    ui.reader.innerHTML = "";
  }
}

function getParcelsForBaque(baqueId) {
  return state.parcels
    .filter((parcel) => parcel.currentBaqueId === baqueId)
    .sort((a, b) => {
      const destinationCompare = getParcelDestinationKey(a).localeCompare(getParcelDestinationKey(b), "fr", { numeric: true });
      if (destinationCompare !== 0) {
        return destinationCompare;
      }
      return new Date(b.updatedAt) - new Date(a.updatedAt);
    });
}

function getBaqueById(baqueId) {
  return state.baques.find((baque) => baque.id === baqueId) || null;
}

function getDestinationRuleById(ruleId) {
  return state.destinationRules.find((rule) => rule.id === ruleId) || null;
}

function getOriginLabel(parcel) {
  return getBaqueById(parcel.originBaqueId)?.name || parcel.originBaqueLabel || "Baque supprimee";
}

function getDestinationGroups() {
  const grouped = state.parcels.reduce((map, parcel) => {
    const group = resolveDestinationGroup(parcel);
    if (!group.key) {
      return map;
    }

    if (!map.has(group.key)) {
      map.set(group.key, {
        key: group.key,
        label: group.label,
        rule: group.rule,
        parcels: [],
        destinations: new Map(),
      });
    }

    const bucket = map.get(group.key);
    const destinationLabel = getParcelDestinationDisplay(parcel);
    bucket.parcels.push(parcel);
    bucket.destinations.set(destinationLabel, (bucket.destinations.get(destinationLabel) || 0) + 1);
    return map;
  }, new Map());

  return [...grouped.values()]
    .map((bucket) => ({
      ...bucket,
      distribution: getDestinationGroupDistribution(bucket.parcels),
      destinations: [...bucket.destinations.entries()]
        .sort(([left], [right]) => left.localeCompare(right, "fr", { numeric: true, sensitivity: "base" })),
    }))
    .sort((left, right) => left.label.localeCompare(right.label, "fr", { numeric: true, sensitivity: "base" }));
}

function resolveDestinationGroup(parcel) {
  const label = getParcelDestinationDisplay(parcel);
  const rule = findMatchingDestinationRule(label);

  if (rule) {
    return {
      key: `rule:${rule.id}`,
      label: rule.label,
      rule,
    };
  }

  return {
    key: `destination:${getParcelDestinationKey(parcel)}`,
    label,
    rule: null,
  };
}

function getDestinationGroupDistribution(parcels) {
  return Object.entries(
    parcels.reduce((acc, parcel) => {
      const baqueName = getBaqueById(parcel.currentBaqueId)?.name || "Baque supprimee";
      acc[baqueName] = (acc[baqueName] || 0) + 1;
      return acc;
    }, {}),
  ).sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0], "fr", { sensitivity: "base" }));
}

function findMatchingDestinationRule(destination) {
  const searchText = normalizeDestinationRuleText(destination);
  if (!searchText) {
    return null;
  }

  return getSortedDestinationRules().find((rule) => doesDestinationRuleMatch(rule, searchText)) || null;
}

function getSortedDestinationRules() {
  return [...state.destinationRules].sort((left, right) => {
    const scoreDifference = getDestinationRuleSpecificity(right) - getDestinationRuleSpecificity(left);
    if (scoreDifference !== 0) {
      return scoreDifference;
    }

    return left.label.localeCompare(right.label, "fr", { sensitivity: "base" });
  });
}

function getDestinationRuleSpecificity(rule) {
  return (rule.patterns.length * 1000) + rule.patterns.join("").length;
}

function doesDestinationRuleMatch(rule, searchText) {
  const matches = rule.patterns.map((pattern) => searchText.includes(pattern));
  return rule.matchMode === "all" ? matches.every(Boolean) : matches.some(Boolean);
}

function countRuleMatchedParcels(rule) {
  return state.parcels.filter(
    (parcel) => findMatchingDestinationRule(getParcelDestinationDisplay(parcel))?.id === rule.id,
  ).length;
}

function getRuleMatchedDestinationLabels(rule) {
  return [...new Set(
    state.parcels
      .map((parcel) => getParcelDestinationDisplay(parcel))
      .filter((destination) => findMatchingDestinationRule(destination)?.id === rule.id),
  )].sort((left, right) => left.localeCompare(right, "fr", { numeric: true, sensitivity: "base" }));
}

function getOrderedBaquesForLayout() {
  return [...state.baques].sort((left, right) => {
    const orderDiff = getBaqueLayoutOrder(left) - getBaqueLayoutOrder(right);
    if (orderDiff !== 0) {
      return orderDiff;
    }

    return left.name.localeCompare(right.name, "fr", { numeric: true, sensitivity: "base" });
  });
}

function getBaqueLayoutOrder(baque) {
  const numericHint = extractBaqueLayoutHint(baque?.name || "");
  if (numericHint !== null) {
    return numericHint;
  }

  const creationIndex = state.baques.findIndex((item) => item.id === baque?.id);
  return creationIndex >= 0 ? creationIndex + 1 : Number.MAX_SAFE_INTEGER;
}

function extractBaqueLayoutHint(value) {
  const match = String(value).match(/(\d{1,3})/);
  return match ? Number.parseInt(match[1], 10) : null;
}

function pickRandomBaque(baques) {
  const list = Array.isArray(baques) ? baques.filter(Boolean) : [];
  if (!list.length) {
    throw new Error("missing-baque");
  }

  return list[Math.floor(Math.random() * list.length)];
}

function normalizeDestinationRule(rule, options = {}) {
  if (!rule) {
    return null;
  }

  const label = normalizeFreeText(String(rule.label || ""));
  const patterns = parseDestinationRulePatterns(rule.patterns || "");
  if (!label || !patterns.length) {
    return null;
  }

  return {
    id: String(rule.id || createId()),
    label,
    matchMode: normalizeDestinationRuleMatchMode(rule.matchMode),
    preferredBaqueId: normalizeDestinationRuleTargetBaqueId(rule.preferredBaqueId || "", options.availableBaqueIds),
    patterns,
    createdAt: normalizeStoredDate(rule.createdAt || "", new Date().toISOString()),
    updatedAt: normalizeStoredDate(rule.updatedAt || rule.createdAt || "", new Date().toISOString()),
  };
}

function normalizeDestinationRuleMatchMode(value) {
  return value === "all" ? "all" : "any";
}

function normalizeDestinationRuleTargetBaqueId(value, availableBaqueIds = null) {
  const baqueId = String(value || "").trim();
  if (!baqueId) {
    return "";
  }

  if (availableBaqueIds instanceof Set) {
    return availableBaqueIds.has(baqueId) ? baqueId : "";
  }

  return hasBaqueId(baqueId) ? baqueId : "";
}

function hasBaqueId(baqueId) {
  return state.baques.some((baque) => baque.id === baqueId);
}

function upsertSimulatedParcel(parcelData, baqueId) {
  const baque = getBaqueById(baqueId);
  if (!baque) {
    return {
      action: "skipped",
      affectedBaqueIds: [],
    };
  }

  const normalizedParcelData = normalizeParcelData(parcelData);
  const commandNumber = getParcelCommandNumber(normalizedParcelData);
  if (!commandNumber && !normalizedParcelData.barcode && !normalizedParcelData.destination && !normalizedParcelData.routeCode) {
    return {
      action: "skipped",
      affectedBaqueIds: [],
    };
  }

  const now = new Date().toISOString();
  const existing = findExistingParcel(state.parcels, normalizedParcelData);
  if (existing) {
    const previousBaqueId = existing.currentBaqueId;
    existing.barcode = normalizedParcelData.barcode;
    existing.commandNumber = commandNumber;
    existing.routeCode = normalizedParcelData.routeCode;
    existing.destination = normalizedParcelData.destination;
    existing.client = normalizedParcelData.client;
    existing.description = normalizedParcelData.description;
    existing.routeLabel = normalizedParcelData.routeLabel;
    existing.reference = normalizedParcelData.reference;
    existing.shippingDate = normalizedParcelData.shippingDate;
    existing.weight = normalizedParcelData.weight;
    existing.measuredDimensions = normalizedParcelData.measuredDimensions;
    existing.packageIndex = normalizedParcelData.packageIndex;
    existing.currentBaqueId = baqueId;
    existing.updatedAt = now;

    return {
      action: "updated",
      affectedBaqueIds: [previousBaqueId, baqueId],
    };
  }

  state.parcels.unshift({
    id: createId(),
    barcode: normalizedParcelData.barcode,
    commandNumber,
    routeCode: normalizedParcelData.routeCode,
    destination: normalizedParcelData.destination,
    client: normalizedParcelData.client,
    description: normalizedParcelData.description,
    routeLabel: normalizedParcelData.routeLabel,
    reference: normalizedParcelData.reference,
    shippingDate: normalizedParcelData.shippingDate,
    weight: normalizedParcelData.weight,
    measuredDimensions: normalizedParcelData.measuredDimensions,
    packageIndex: normalizedParcelData.packageIndex,
    currentBaqueId: baqueId,
    originBaqueId: baqueId,
    originBaqueLabel: baque.name,
    createdAt: now,
    updatedAt: now,
  });

  return {
    action: "created",
    affectedBaqueIds: [baqueId],
  };
}

function getParcelDestinationDisplay(parcel) {
  return parcel.destination || formatRouteCodeForDisplay(parcel.routeCode) || "Sans destination";
}

function getParcelDestinationKey(parcel) {
  return getParcelDestinationDisplay(parcel).toUpperCase();
}

function getDestinationRuleMatchModeLabel(matchMode) {
  return matchMode === "all" ? "Tous les mots-cles" : "Au moins un mot-cle";
}

function renderRouteCodeMeta(parcels) {
  const routeCodes = [...new Set(parcels.map((parcel) => parcel.routeCode).filter(Boolean))];
  if (routeCodes.length !== 1) {
    return "";
  }

  return `<span><strong>Route :</strong> ${escapeHtml(formatRouteCodeForDisplay(routeCodes[0]))}</span>`;
}

async function getPdfJs() {
  ensurePdfJsCompatibility();
  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    return window.pdfjsLib;
  }

  if (!pdfjsLibPromise) {
    pdfjsLibPromise = loadPdfJsScript()
      .then(() => {
        if (!window.pdfjsLib) {
          throw new Error("pdfjs-unavailable");
        }

        window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
        return window.pdfjsLib;
      })
      .catch((error) => {
        pdfjsLibPromise = null;
        throw error;
      });
  }

  return pdfjsLibPromise;
}

function loadPdfJsScript() {
  if (window.pdfjsLib) {
    return Promise.resolve(window.pdfjsLib);
  }

  return new Promise((resolve, reject) => {
    const existingScript = document.querySelector('script[data-role="pdfjs-runtime"]');
    if (existingScript instanceof HTMLScriptElement) {
      existingScript.addEventListener("load", () => resolve(window.pdfjsLib), { once: true });
      existingScript.addEventListener("error", () => reject(new Error("pdfjs-script-load-failed")), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = PDFJS_SCRIPT_URL;
    script.defer = true;
    script.dataset.role = "pdfjs-runtime";
    script.onload = () => resolve(window.pdfjsLib);
    script.onerror = () => reject(new Error("pdfjs-script-load-failed"));
    document.head.append(script);
  });
}

function ensurePdfJsCompatibility() {
  if (typeof Promise.withResolvers === "function") {
    return;
  }

  Promise.withResolvers = function withResolvers() {
    let resolve;
    let reject;
    const promise = new Promise((promiseResolve, promiseReject) => {
      resolve = promiseResolve;
      reject = promiseReject;
    });

    return { promise, resolve, reject };
  };
}

async function renderPdfPageToCanvas(page) {
  const viewport = page.getViewport({ scale: 3 });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { alpha: false });

  if (!context) {
    throw new Error("canvas-unavailable");
  }

  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);

  await page.render({
    canvasContext: context,
    viewport,
  }).promise;

  return await canvasToBlob(canvas);
}

function canvasToBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (blob) {
          resolve(blob);
        } else {
          reject(new Error("canvas-blob-failed"));
        }
      },
      "image/png",
      1,
    );
  });
}

function openPdfDatabase() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error("indexeddb-unavailable"));
      return;
    }

    const request = window.indexedDB.open(PDF_DB_NAME, 1);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(PDF_STORE_NAME)) {
        db.createObjectStore(PDF_STORE_NAME);
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("indexeddb-open-failed"));
  });
}

async function getDeliveryNoteFile(noteId) {
  const db = await openPdfDatabase();

  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction(PDF_STORE_NAME, "readonly");
      const request = transaction.objectStore(PDF_STORE_NAME).get(noteId);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("indexeddb-read-failed"));
    });
  } finally {
    db.close();
  }
}

async function saveDeliveryNoteFile(noteId, file) {
  const db = await openPdfDatabase();

  try {
    await new Promise((resolve, reject) => {
      const transaction = db.transaction(PDF_STORE_NAME, "readwrite");
      transaction.objectStore(PDF_STORE_NAME).put(file, noteId);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("indexeddb-write-failed"));
      transaction.onabort = () => reject(transaction.error || new Error("indexeddb-write-aborted"));
    });
  } finally {
    db.close();
  }
}

async function deleteDeliveryNoteFile(noteId) {
  const db = await openPdfDatabase();

  try {
    await new Promise((resolve, reject) => {
      const transaction = db.transaction(PDF_STORE_NAME, "readwrite");
      transaction.objectStore(PDF_STORE_NAME).delete(noteId);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error || new Error("indexeddb-delete-failed"));
      transaction.onabort = () => reject(transaction.error || new Error("indexeddb-delete-aborted"));
    });
  } finally {
    db.close();
  }
}

function showToast(message, type = "default") {
  const toast = document.createElement("div");
  toast.className = `toast${type === "danger" ? " toast--danger" : ""}`;
  toast.textContent = message;
  ui.toastZone.appendChild(toast);

  window.setTimeout(() => {
    toast.remove();
  }, 3200);
}
