const STORAGE_KEY = "transbaus-gaviota-state-v1";
const COLLAPSE_STORAGE_KEY = "le-baus-du-tri-collapse-v1";
const ACCESS_STORAGE_KEY = "transbaus-gaviota-access-v1";
const ACCESS_RATE_LIMIT_STORAGE_KEY = "transbaus-gaviota-access-rate-v1";
const ACCESS_PASSWORD = "2005";
const ACCESS_FAILED_ATTEMPTS_LIMIT = 3;
const ACCESS_LOCK_DURATION_MS = 10_000;
const PDF_DB_NAME = "le-baus-du-tri-documents-v1";
const PDF_STORE_NAME = "delivery-notes";
const PDFJS_SCRIPT_URL = "vendor/pdf.min.js";
const PDFJS_WORKER_URL = "vendor/pdf.worker.min.js";
const DEFAULT_COLLAPSE_STATE = {
  flow: true,
  scanner: false,
  baqueForm: true,
  search: true,
  deliveryNote: true,
  destinations: true,
  baques: true,
};
const DEFAULT_BAQUES = [
  { name: "Baque 1", location: "Zone A" },
  { name: "Baque 2", location: "Zone B" },
  { name: "Baque 3", location: "Zone C" },
  { name: "Baque 4", location: "Zone D" },
];

const state = loadState();
const collapseState = loadCollapseState();
const accessRateLimit = loadAccessRateLimit();
const ui = {};
const scanner = {
  instance: null,
  active: false,
  handled: false,
  importingBarcode: false,
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

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  render();
  syncAccessGate();
});

function cacheElements() {
  ui.loginGate = document.querySelector("#loginGate");
  ui.loginForm = document.querySelector("#loginForm");
  ui.loginPasswordInput = document.querySelector("#loginPasswordInput");
  ui.loginStatus = document.querySelector("#loginStatus");
  ui.loginSubmitBtn = ui.loginForm?.querySelector('button[type="submit"]');
  ui.logoutBtn = document.querySelector("#logoutBtn");
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
  ui.packageIndexInput = document.querySelector("#packageIndexInput");
  ui.barcodeInput = document.querySelector("#barcodeInput");
  ui.openScannerBtn = document.querySelector("#openScannerBtn");
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
  ui.scannerStatus = document.querySelector("#scannerStatus");
  ui.closeScannerBtn = document.querySelector("#closeScannerBtn");
  ui.captureModal = document.querySelector("#captureModal");
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
  ui.logoutBtn.addEventListener("click", handleLogoutClick);
  ui.parcelForm.addEventListener("submit", handleParcelSubmit);
  ui.baqueForm.addEventListener("submit", handleBaqueSubmit);
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
  ui.openScannerBtn.addEventListener("click", openScanner);
  ui.importBarcodeBtn.addEventListener("click", openBarcodeCameraPicker);
  ui.chooseBarcodeBtn.addEventListener("click", openBarcodeLibraryPicker);
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
  ui.baquesGrid.addEventListener("click", handleBaqueGridClick);
  ui.baquesGrid.addEventListener("change", handleBaqueGridChange);
  window.addEventListener("beforeunload", () => {
    void stopScanner();
    void stopCaptureStream();
    void stopOcrWorker();
  });
  window.addEventListener("resize", applyCollapseStateToDom);
}

function syncAccessGate() {
  refreshAccessRateLimit();
  setAppAccess(hasStoredAccess());
}

function handleLoginSubmit(event) {
  event.preventDefault();

  refreshAccessRateLimit();
  if (isAccessTemporarilyLocked()) {
    syncAccessRateLimitUi();
    return;
  }

  const typedPassword = ui.loginPasswordInput.value.trim();
  if (typedPassword !== ACCESS_PASSWORD) {
    registerFailedLoginAttempt();
    if (!isAccessTemporarilyLocked()) {
      ui.loginPasswordInput.focus();
      ui.loginPasswordInput.select();
    }
    return;
  }

  resetAccessRateLimit();
  window.localStorage.setItem(ACCESS_STORAGE_KEY, "granted");
  ui.loginStatus.textContent = "";
  ui.loginForm.reset();
  setAppAccess(true);
}

function handleLogoutClick() {
  window.localStorage.removeItem(ACCESS_STORAGE_KEY);
  void closeScanner();
  void closeCaptureModal();
  setAppAccess(false);
}

function hasStoredAccess() {
  return window.localStorage.getItem(ACCESS_STORAGE_KEY) === "granted";
}

function setAppAccess(isGranted) {
  document.body.classList.toggle("app-locked", !isGranted);
  ui.loginGate.setAttribute("aria-hidden", String(isGranted));
  ui.logoutBtn.hidden = !isGranted;

  if (!isGranted) {
    ui.loginForm.reset();
    syncAccessRateLimitUi();
    if (!isAccessTemporarilyLocked()) {
      ui.loginPasswordInput.focus();
    }
    return;
  }

  stopAccessLockCountdown();
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

function loadState() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return createDefaultState();
    }

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed.baques) || !Array.isArray(parsed.parcels)) {
      return createDefaultState();
    }

    const baques = parsed.baques.map((baque) => ({
      id: baque.id || createId(),
      name: String(baque.name || "Baque"),
      location: String(baque.location || "Sans emplacement"),
      validatedAt: normalizeStoredDate(baque.validatedAt || "", ""),
      createdAt: normalizeStoredDate(baque.createdAt, new Date().toISOString()),
    }));

    const knownIds = new Set(baques.map((baque) => baque.id));
    const parcels = parsed.parcels
      .filter((parcel) => parcel && knownIds.has(parcel.currentBaqueId))
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
        .map((rule) => normalizeDestinationRule(rule))
        .filter(Boolean)
      : [];

    return {
      baques: baques.length ? baques : createDefaultState().baques,
      parcels,
      deliveryNotes,
      destinationRules,
    };
  } catch (error) {
    return createDefaultState();
  }
}

function createDefaultState() {
  return {
    baques: DEFAULT_BAQUES.map((baque) => ({
      id: createId(),
      name: baque.name,
      location: baque.location,
      validatedAt: "",
      createdAt: new Date().toISOString(),
    })),
    parcels: [],
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
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function render() {
  renderHeroStats();
  renderBaqueSelect();
  renderDestinationRuleTargetOptions();
  renderDestinationRules();
  renderDestinationSummary();
  renderSortingPlan();
  renderBaques();
  renderSearchResults();
  renderDeliveryNotes();
  applyCollapseStateToDom();
}

function renderHeroStats() {
  const totalBaques = state.baques.length;
  const totalParcels = state.parcels.length;
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
              <h3>${escapeHtml(group.label)}</h3>
              ${group.rule ? `<span class="tag">Regle</span>` : ""}
            </div>
            ${quickAction}
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
  const detailLines = [
    `Destination <strong>${escapeHtml(displayDestination)}</strong>`,
    parcel.commandNumber ? `Numero de commande : ${escapeHtml(parcel.commandNumber)}` : "",
    parcel.barcode && parcel.barcode !== parcel.commandNumber ? `Code-barres : ${escapeHtml(parcel.barcode)}` : "",
    parcel.client ? `Client : ${escapeHtml(parcel.client)}` : "",
    parcel.routeCode ? `Numero destination : ${escapeHtml(displayRouteCode)}` : "",
    parcel.routeLabel ? `Route : ${escapeHtml(parcel.routeLabel)}` : "",
    parcel.reference ? `Reference : ${escapeHtml(parcel.reference)}` : "",
    parcel.packageIndex ? `Colis : ${escapeHtml(parcel.packageIndex)}` : "",
    parcel.weight ? `Poids : ${escapeHtml(parcel.weight)}` : "",
    parcel.shippingDate ? `Date : ${escapeHtml(parcel.shippingDate)}` : "",
    `Origine : ${escapeHtml(getOriginLabel(parcel))}`,
    `Derniere mise a jour : ${escapeHtml(formatDate(parcel.updatedAt || parcel.createdAt))}`,
  ]
    .filter(Boolean)
    .join("<br>");
  const tagLabel = displayRouteCode || getDestinationShortLabel(displayDestination) || "Colis";
  const parcelHeading = getParcelIdentifier(parcel);

  return `
    <article class="parcel-item" data-parcel-id="${escapeHtml(parcel.id)}">
      <div class="parcel-item__top">
        <div>
          <p class="parcel-code">${escapeHtml(parcelHeading)}</p>
          <p class="parcel-meta">${detailLines}</p>
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

function handleBaqueSubmit(event) {
  event.preventDefault();

  const name = ui.baqueNameInput.value.trim();
  const location = ui.baqueLocationInput.value.trim();

  if (!name || !location) {
    showToast("Le nom et l'emplacement de la baque sont obligatoires.", "danger");
    return;
  }

  state.baques.push({
    id: createId(),
    name,
    location,
    validatedAt: "",
    createdAt: new Date().toISOString(),
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

  state.destinationRules.push({
    id: createId(),
    label,
    matchMode,
    preferredBaqueId,
    patterns,
    createdAt: new Date().toISOString(),
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
  saveState();
  render();
  showToast("Baque mise a jour.");
}

function handleModalClick(event) {
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
  const isMobile = window.matchMedia("(max-width: 760px)").matches;

  document.querySelectorAll("[data-collapsible-key]").forEach((section) => {
    if (!(section instanceof HTMLElement)) {
      return;
    }

    const sectionKey = section.dataset.collapsibleKey;
    if (!sectionKey) {
      return;
    }

    const shouldCollapse = isMobile && Boolean(collapseState[sectionKey]);
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

  const deliveryNote = {
    id: createId(),
    name: normalizeFreeText(file.name || "Bon-de-livraison.pdf"),
    size: Number(file.size || 0),
    importedAt: new Date().toISOString(),
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
    const analysis = compareDeliveryNoteEntries(entries);

    deliveryNote.analysis = {
      ...analysis,
      analyzedAt: new Date().toISOString(),
    };

    saveState();
    renderDeliveryNotes();
    ui.deliveryNoteStatus.textContent = analysis.totalMissingCount
      ? `${analysis.totalMissingCount} colis manquants identifies dans ${deliveryNote.name}.`
      : analysis.parseError
      ? `Analyse terminee, mais aucune livraison exploitable n'a ete detectee dans ${deliveryNote.name}.`
      : `Aucun colis manquant detecte dans ${deliveryNote.name}.`;
  } catch (error) {
    const errorMessage = getDeliveryNoteErrorMessage(error);
    deliveryNote.analysis = {
      totalEntries: 0,
      totalExpectedCount: 0,
      totalRegisteredCount: 0,
      totalMissingCount: 0,
      incomparableParcelsCount: countIncomparableParcels(),
      parseError: errorMessage,
      missingEntries: [],
      analyzedAt: new Date().toISOString(),
    };
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

    const analysis = compareDeliveryNoteEntries(entries);
    deliveryNote.analysis = {
      ...analysis,
      analyzedAt: new Date().toISOString(),
    };

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
    ui.barcodeStatus.textContent = "Analyse de la photo du code-barres...";

    fileScanner = new Html5Qrcode("barcodeFileReader");
    const decodedText = await fileScanner.scanFile(file, false);
    const normalizedCode = decodedText.trim();

    ui.barcodeInput.value = normalizedCode;

    const added = upsertParcel(normalizedCode);
    if (!added) {
      ui.barcodeStatus.textContent = "Code-barres detecte. Verifiez le numero destination puis enregistrez.";
      showToast("Code-barres detecte.");
    } else {
      ui.barcodeStatus.textContent = "Code-barres detecte et applique au colis.";
    }
  } catch (error) {
    ui.barcodeStatus.textContent = "Impossible de lire le code-barres sur cette photo.";
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
  if (!file) {
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
    const result = await worker.recognize(file);
    const parsed = parseLabelText(result.data.text);

    applyParsedLabelData(parsed);

    if (!parsed.routeCode) {
      ui.ocrStatus.textContent = "Lecture terminee, mais peu d'informations ont ete reconnues. Reprenez une photo plus nette.";
      showToast("Numero destination introuvable. Essayez une photo plus nette.", "danger");
      return;
    }

    ui.ocrStatus.textContent = "Etiquette analysee. Verifiez les champs puis enregistrez le colis.";
    showToast("Etiquette analysee. Les informations ont ete remplies.");
  } catch (error) {
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
      ? "Cadrez l'etiquette dans le rectangle, puis prenez la photo."
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
    ? "Placez l'etiquette entiere dans le cadre. La photo se prend automatiquement quand elle est stable."
    : "Placez le code-barres au centre du cadre, evitez les reflets, puis prenez la photo.";
  ui.takeCaptureBtn.textContent = isLabelMode ? "Prendre maintenant" : "Prendre la photo";
  ui.captureGuide.classList.toggle("capture-guide--label", isLabelMode);
  ui.captureGuide.classList.toggle("capture-guide--barcode", !isLabelMode);
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
    const file = await captureCurrentFrame(mode);
    await closeCaptureModal({ force: true, silent: true });

    if (mode === "label") {
      await processLabelFile(file);
    } else {
      await processBarcodeFile(file);
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
  ui.captureStatus.textContent = "Auto actif. Placez l'etiquette dans le cadre et gardez le telephone stable.";

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
      ui.captureStatus.textContent = analysis.message;
      return;
    }

    captureSession.stableFrameCount += 1;

    if (captureSession.stableFrameCount < 3) {
      ui.captureStatus.textContent = "Etiquette detectee. Restez immobile une seconde...";
      return;
    }

    captureSession.autoTriggered = true;
    ui.captureStatus.textContent = "Etiquette stable detectee. Photo automatique...";
    void handleCapturePhoto({ auto: true });
  }, 420);
}

function stopAutoCaptureMonitoring() {
  if (captureSession.autoCaptureTimer) {
    window.clearInterval(captureSession.autoCaptureTimer);
    captureSession.autoCaptureTimer = 0;
  }

  captureSession.lastFrameSignature = null;
  captureSession.stableFrameCount = 0;
  captureSession.autoTriggered = false;
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
  const isStable = motion !== null && motion <= 11;
  const isReady = whiteRatio >= 0.34 && centerWhiteRatio >= 0.52 && contrast >= 34 && edgeRatio >= 0.06 && isStable;

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
  if (metrics.motion === null || metrics.motion > 11) {
    return "Auto actif. Restez immobile et gardez l'etiquette bien droite dans le cadre.";
  }

  if (metrics.centerWhiteRatio < 0.52 || metrics.whiteRatio < 0.34) {
    return "Auto actif. Rapprochez ou recentrez l'etiquette pour qu'elle remplisse mieux le cadre.";
  }

  if (metrics.contrast < 34 || metrics.edgeRatio < 0.06) {
    return "Auto actif. Evitez le flou et les reflets sur l'etiquette.";
  }

  return "Etiquette detectee. Restez immobile...";
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

  const extraMargin = mode === "label" ? 0.04 : 0.1;
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

function parseDeliveryNoteText(text) {
  const lines = String(text)
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => normalizePdfOcrLine(line))
    .filter(Boolean);
  const structuredEntries = dedupeDeliveryEntries(parseStructuredDeliveryNoteLines(lines));
  const legacyEntries = dedupeDeliveryEntries(parseLegacyDeliveryNoteLines(lines));

  return structuredEntries.length >= legacyEntries.length ? structuredEntries : legacyEntries;
}

function parseLegacyDeliveryNoteLines(lines) {
  const entries = [];

  lines.forEach((line, index) => {
    const rawCommandNumber = extractDeliveryCommandNumber(line);
    if (!rawCommandNumber) {
      return;
    }

    const contextLines = lines.slice(Math.max(0, index - 6), Math.min(lines.length, index + 6));
    const commandNumber = resolveDeliveryCommandNumber(rawCommandNumber, contextLines);
    const expectedCount = extractDeliveryPackageCount(contextLines, index - Math.max(0, index - 6));
    entries.push({
      commandNumber,
      expectedCount,
      client: extractDeliveryClient(contextLines),
      city: extractDeliveryCity(contextLines),
      rawContext: contextLines.join(" | "),
    });
  });

  return entries;
}

function compareDeliveryNoteEntries(entries) {
  const incomparableParcelsCount = countIncomparableParcels();
  if (!entries.length) {
    return {
      totalEntries: 0,
      totalExpectedCount: 0,
      totalRegisteredCount: 0,
      totalMissingCount: 0,
      incomparableParcelsCount,
      missingEntries: [],
      parseError: "Aucune livraison exploitable n'a ete detectee dans ce PDF. Le scan est peut-etre trop flou.",
    };
  }

  const registeredCommandCounts = buildRegisteredCommandCounts();
  const registeredCommandInfo = buildRegisteredCommandInfo();
  const missingEntries = [];
  let totalExpectedCount = 0;
  let totalRegisteredCount = 0;

  entries.forEach((entry) => {
    const expectedCount = Math.max(1, Number(entry.expectedCount || 1));
    const registeredCount = registeredCommandCounts.get(entry.commandNumber) || 0;
    const matchedCount = Math.min(expectedCount, registeredCount);
    const missingCount = Math.max(0, expectedCount - registeredCount);

    totalExpectedCount += expectedCount;
    totalRegisteredCount += matchedCount;

    if (missingCount > 0) {
      const registeredInfo = registeredCommandInfo.get(entry.commandNumber);
      missingEntries.push({
        ...entry,
        client: registeredInfo?.client || entry.client,
        city: registeredInfo?.city || entry.city,
        expectedCount,
        registeredCount,
        missingCount,
      });
    }
  });

  return {
    totalEntries: entries.length,
    totalExpectedCount,
    totalRegisteredCount,
    totalMissingCount: Math.max(0, totalExpectedCount - totalRegisteredCount),
    incomparableParcelsCount,
    missingEntries,
    parseError: "",
  };
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
  ui.importBarcodeBtn.disabled = isBusy;
  ui.chooseBarcodeBtn.disabled = isBusy;
  ui.importBarcodeBtn.textContent = isBusy ? "Analyse en cours..." : "Prendre une photo";
  ui.chooseBarcodeBtn.textContent = isBusy ? "Analyse en cours..." : "Choisir une photo";
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

  if (parsed.packageIndex) {
    ui.packageIndexInput.value = parsed.packageIndex;
  }

  if (parsed.barcode) {
    ui.barcodeInput.value = parsed.barcode;
  }
}

function parseLabelText(text) {
  const lines = text
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => normalizeOcrLine(line))
    .filter(Boolean);
  const sections = extractLabelSections(lines);
  const parsed = {
    rawText: text,
    barcode: extractCommandeNumber(lines, text),
    commandNumber: extractCommandeNumber(lines, text),
    routeCode: extractRouteCode(lines, text),
    destination: extractDestination(lines, sections),
    client: sections.client,
    description: sections.description,
    routeLabel: sections.route,
    reference: sections.reference,
    shippingDate: extractShippingDate(text),
    weight: extractWeight(text),
    packageIndex: extractPackageIndex(lines, text),
  };

  return normalizeParcelData(parsed);
}

function extractLabelSections(lines) {
  const sections = {
    client: [],
    address: [],
    description: [],
    route: [],
    reference: [],
  };
  let currentSection = "";

  for (const line of lines) {
    const match = matchSectionHeader(line);
    if (match) {
      currentSection = match.section;
      if (match.inlineValue) {
        sections[currentSection].push(match.inlineValue);
      }
      continue;
    }

    if (!currentSection) {
      continue;
    }

    if (looksLikeMetaLine(line)) {
      continue;
    }

    sections[currentSection].push(line);
  }

  return {
    client: joinSectionLines(sections.client),
    address: joinSectionLines(sections.address, ", "),
    description: joinSectionLines(sections.description),
    route: joinSectionLines(sections.route),
    reference: joinSectionLines(sections.reference),
  };
}

function matchSectionHeader(line) {
  const patterns = [
    { section: "client", regex: /^CLIENT\b[:\s-]*(.*)$/i },
    { section: "address", regex: /^ADRESSE\b[:\s-]*(.*)$/i },
    { section: "description", regex: /^DESCRIPTION\b[:\s-]*(.*)$/i },
    { section: "route", regex: /^ROUTE\b[:\s-]*(.*)$/i },
    { section: "reference", regex: /^REF\b[:\s-]*(.*)$/i },
  ];

  for (const pattern of patterns) {
    const result = line.match(pattern.regex);
    if (result) {
      return {
        section: pattern.section,
        inlineValue: normalizeFreeText(result[1] || ""),
      };
    }
  }

  return null;
}

function looksLikeMetaLine(line) {
  return [
    /GAVIOTA/i,
    /SALEILLES/i,
    /C\.?I\.?F/i,
    /^TEL/i,
    /COMMANDE/i,
    /^DATE/i,
    /^\d+\s*\/\s*\d+$/,
    /^\d+[.,]\d+\s*KG$/i,
    /^R\s*(?:\d\s*){5,7}$/i,
  ].some((regex) => regex.test(line));
}

function extractCommandeNumber(lines, text) {
  const commandIndex = lines.findIndex((line) => /COMMANDE/i.test(line));
  if (commandIndex >= 0) {
    const nearbyLines = lines.slice(commandIndex, commandIndex + 5);
    const inlineMatch = nearbyLines[0]?.match(/COMMANDE[^\dA-Z]*([0-9]{5,10})/i);
    if (inlineMatch) {
      return inlineMatch[1].replace(/\s+/g, "");
    }

    const candidates = nearbyLines
      .flatMap((line, offset) =>
        [...line.matchAll(/\b\d{5,10}\b/g)].map((match) => ({
          value: match[0],
          line,
          offset,
        })),
      )
      .filter((candidate, index, array) =>
        array.findIndex((item) => item.value === candidate.value) === index,
      )
      .sort((left, right) => getCommandeCandidateScore(right) - getCommandeCandidateScore(left));

    if (candidates[0]) {
      return candidates[0].value.replace(/\s+/g, "");
    }
  }

  const directMatch = text.match(/(?:N\W*COMMANDE|COMMANDE)[^\dA-Z]*([0-9]{5,10})/i);
  if (directMatch) {
    return directMatch[1].replace(/\s+/g, "");
  }

  return "";
}

function getCommandeCandidateScore(candidate) {
  const normalizedLine = normalizeFreeText(candidate.line || "");
  const normalizedValue = String(candidate.value || "").replace(/\s+/g, "");
  let score = 0;

  if (/COMMANDE/i.test(normalizedLine)) {
    score += 50;
  }

  if (new RegExp(`^${normalizedValue}$`).test(normalizedLine)) {
    score += 18;
  }

  if (normalizedValue.length >= 6 && normalizedValue.length <= 8) {
    score += 14;
  } else if (normalizedValue.length === 5) {
    score += 4;
  }

  if (normalizedValue.startsWith("0")) {
    score += 6;
  }

  score -= candidate.offset * 4;

  return score;
}

function extractRouteCode(lines, text) {
  const routeCodeRegex = /\bR\s*(?:\d\s*){5,7}\b/i;
  const fromLines = lines.find((line) => routeCodeRegex.test(line));
  if (fromLines) {
    return normalizeRouteCode(fromLines.match(routeCodeRegex)[0]);
  }

  const fromText = text.match(routeCodeRegex);
  return fromText ? normalizeRouteCode(fromText[0]) : "";
}

function extractDestination(lines, sections) {
  if (sections.address) {
    return sections.address;
  }

  const postalLine = lines.find((line) => /\b\d{5}\b/.test(line) && !/SALEILLES/i.test(line));
  return postalLine ? sanitizeDestination(postalLine) : "";
}

function extractShippingDate(text) {
  const match = text.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
  return match ? match[0] : "";
}

function extractWeight(text) {
  const match = text.match(/\b\d{1,3}[.,]\d{1,2}\s*Kg\b/i);
  return match ? normalizeFreeText(match[0].replace(/\s+/g, " ")) : "";
}

function extractPackageIndex(lines, text) {
  const lineMatch = lines.find((line) => /^\d{1,2}\s*\/\s*\d{1,2}$/.test(line));
  if (lineMatch) {
    return lineMatch.replace(/\s+/g, "");
  }

  const match = text.match(/\b\d{1,2}\s*\/\s*\d{1,2}\b(?!\s*\/)/);
  return match ? match[0].replace(/\s+/g, "") : "";
}

function joinSectionLines(lines, separator = " ") {
  return normalizeFreeText(lines.filter(Boolean).join(separator));
}

function normalizeOcrLine(value) {
  return normalizeFreeText(
    value
      .replace(/[|]/g, "I")
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'"),
  );
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
  const existing = findExistingParcel(normalizedParcelData);
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
  ui.packageIndexInput.value = "";
  ui.barcodeInput.value = "";
  ui.ocrStatus.textContent = "";
  ui.barcodeStatus.textContent = "";
  ui.routeCodeInput.focus();
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

  state.baques = state.baques.filter((item) => item.id !== baqueId);
  state.parcels = state.parcels.filter((parcel) => parcel.currentBaqueId !== baqueId);
  state.destinationRules = state.destinationRules.map((rule) => (
    rule.preferredBaqueId === baqueId
      ? { ...rule, preferredBaqueId: "" }
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
    saveState();
    render();
    showToast(`Validation retiree pour ${baque.name}.`);
    return;
  }

  baque.validatedAt = new Date().toISOString();
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

async function openScanner() {
  if (typeof window.Html5QrcodeScanner === "undefined") {
    showToast("La librairie de scan n'a pas pu etre chargee.", "danger");
    return;
  }

  if (scanner.active) {
    ui.scannerModal.classList.remove("hidden");
    ui.scannerModal.setAttribute("aria-hidden", "false");
    return;
  }

  scanner.handled = false;
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
        ui.barcodeInput.value = decodedText.trim();
        const added = upsertParcel(decodedText.trim());
        if (!added) {
          showToast("Code detecte. Completez les champs puis validez.");
        }

        await closeScanner();
      },
      () => {},
    );

    scanner.active = true;
    ui.scannerStatus.textContent = "Scanner actif. Vous pouvez utiliser la camera ou importer une photo du code.";
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

function createId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `id-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
}

function pickRandomBaque(baques) {
  const list = Array.isArray(baques) ? baques.filter(Boolean) : [];
  if (!list.length) {
    throw new Error("missing-baque");
  }

  return list[Math.floor(Math.random() * list.length)];
}

function normalizeDestination(value) {
  return sanitizeDestination(value);
}

function normalizeRouteCode(value) {
  return normalizeFreeText(value).replace(/\s+/g, "").toUpperCase();
}

function normalizeFreeText(value) {
  return value.trim().replace(/\s+/g, " ");
}

function normalizeDestinationRule(rule) {
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
    preferredBaqueId: normalizeDestinationRuleTargetBaqueId(rule.preferredBaqueId || ""),
    patterns,
    createdAt: rule.createdAt || new Date().toISOString(),
  };
}

function normalizeDestinationRuleMatchMode(value) {
  return value === "all" ? "all" : "any";
}

function normalizeDestinationRuleTargetBaqueId(value) {
  const baqueId = String(value || "").trim();
  return hasBaqueId(baqueId) ? baqueId : "";
}

function hasBaqueId(baqueId) {
  return state.baques.some((baque) => baque.id === baqueId);
}

function parseDestinationRulePatterns(value) {
  const rawValues = Array.isArray(value)
    ? value
    : String(value || "")
      .replace(/,/g, "\n")
      .split("\n");

  return [...new Set(rawValues.map((entry) => normalizeDestinationRulePattern(entry)).filter(Boolean))];
}

function normalizeDestinationRulePattern(value) {
  const rawValue = normalizeFreeText(String(value || ""));
  const directPostalCode = rawValue.match(/^\d{5}$/)?.[0];
  if (directPostalCode) {
    return directPostalCode;
  }

  return normalizeDestinationRuleText(rawValue);
}

function normalizeDestinationRuleText(value) {
  return stripDiacritics(sanitizeDestination(String(value || "")).toUpperCase())
    .replace(/[^0-9A-Z\s-]/g, " ")
    .replace(/[-/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stripDiacritics(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeParcelData(parcel) {
  const barcode = normalizeBarcode(parcel.barcode || "");
  const explicitCommandNumber = normalizeCommandNumber(parcel.commandNumber || "");
  const routeLabel = normalizeFreeText(parcel.routeLabel || "");
  const rawDestination = sanitizeDestination(parcel.destination || "");
  const shippingDate = normalizeFreeText(parcel.shippingDate || "");
  const routeCode = reconcileRouteCode(parcel.routeCode || "", routeLabel, rawDestination);
  const destination = reconcileDestination(rawDestination, routeCode);
  const packageIndex = sanitizePackageIndex(parcel.packageIndex || "", shippingDate);
  const normalizedBarcode = reconcileBarcode(barcode, destination, routeCode);
  const commandNumber = reconcileCommandNumber(explicitCommandNumber, normalizedBarcode);

  return {
    ...parcel,
    barcode: normalizedBarcode,
    commandNumber,
    routeCode,
    destination,
    client: normalizeFreeText(parcel.client || ""),
    description: normalizeFreeText(parcel.description || ""),
    routeLabel,
    reference: normalizeFreeText(parcel.reference || ""),
    shippingDate,
    weight: normalizeFreeText(parcel.weight || ""),
    packageIndex,
  };
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
  const existing = findExistingParcel(normalizedParcelData);
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

function getParcelIdentifier(parcel) {
  return getParcelCommandNumber(parcel) || parcel.barcode || formatRouteCodeForDisplay(parcel.routeCode) || "Sans code";
}

function buildDeliveryEntryDestinationLabel(entry) {
  const city = sanitizeDestination(entry?.city || "");
  const postalCode = extractDeliveryEntryPostalCode(entry);
  if (postalCode && city) {
    return `${postalCode} ${city}`;
  }

  return city || postalCode || "";
}

function extractDeliveryEntryPostalCode(entry) {
  const rawContext = normalizeDeliveryTextLine(entry?.rawContext || "");
  const match = rawContext.match(/\b\d{5}\b/);
  return match ? match[0] : "";
}

function findExistingParcel(parcelData) {
  const barcode = normalizeBarcode(parcelData.barcode || "");
  const packageIndex = normalizeFreeText(parcelData.packageIndex || "");
  const commandNumber = getParcelCommandNumber(parcelData);

  if (commandNumber && packageIndex) {
    const exactCommandMatch = state.parcels.find(
      (parcel) => getParcelCommandNumber(parcel) === commandNumber && normalizeFreeText(parcel.packageIndex || "") === packageIndex,
    );
    if (exactCommandMatch) {
      return exactCommandMatch;
    }
  }

  if (barcode && packageIndex) {
    const exactBarcodeMatch = state.parcels.find(
      (parcel) => normalizeBarcode(parcel.barcode || "") === barcode && normalizeFreeText(parcel.packageIndex || "") === packageIndex,
    );
    if (exactBarcodeMatch) {
      return exactBarcodeMatch;
    }
  }

  if (commandNumber && !packageIndex) {
    const commandMatches = state.parcels.filter(
      (parcel) => getParcelCommandNumber(parcel) === commandNumber && !normalizeFreeText(parcel.packageIndex || ""),
    );
    if (commandMatches.length === 1) {
      return commandMatches[0];
    }
  }

  if (barcode && !packageIndex) {
    const barcodeMatches = state.parcels.filter(
      (parcel) => normalizeBarcode(parcel.barcode || "") === barcode && !normalizeFreeText(parcel.packageIndex || ""),
    );
    if (barcodeMatches.length === 1) {
      return barcodeMatches[0];
    }
  }

  return null;
}

function sanitizeDestination(value) {
  const rawSegments = String(value)
    .replaceAll(",", "\n")
    .split("\n")
    .map((segment) => normalizeFreeText(segment))
    .filter(Boolean);

  const cleanedSegments = rawSegments
    .map((segment) => cleanDestinationSegment(segment))
    .filter(Boolean);

  const postalIndex = cleanedSegments.findIndex((segment) => /\b\d{5}\b/.test(segment));
  const keptSegments = postalIndex >= 0 ? cleanedSegments.slice(postalIndex) : cleanedSegments;

  return normalizeFreeText(keptSegments.join(" "));
}

function cleanDestinationSegment(segment) {
  let nextSegment = normalizeFreeText(segment);
  const postalMatch = nextSegment.match(/\b\d{5}\b/);
  if (postalMatch) {
    nextSegment = nextSegment.slice(postalMatch.index).trim();
  }

  nextSegment = nextSegment
    .replace(/^[^0-9A-ZÀ-ÿ-]+/u, "")
    .replace(/\b[Il|]{1,3}\b/g, "")
    .replace(/\s+\d{1,2}$/u, "")
    .replace(/\s{2,}/g, " ")
    .trim();

  if (!nextSegment) {
    return "";
  }

  if (/^\d+$/.test(nextSegment) || /^[Il|]+$/i.test(nextSegment)) {
    return "";
  }

  return nextSegment;
}

function sanitizePackageIndex(packageIndex, shippingDate) {
  const normalized = normalizeFreeText(packageIndex).replace(/\s+/g, "");
  if (!normalized) {
    return "";
  }

  if (!/^\d{1,2}\/\d{1,2}$/.test(normalized)) {
    return "";
  }

  const datePrefix = shippingDate.match(/^(\d{2}\/\d{2})\//)?.[1];
  if (datePrefix && normalized === datePrefix) {
    return "";
  }

  return normalized;
}

function reconcileDestination(destination, routeCode) {
  const routePostalCode = routeCode.match(/(\d{5})$/)?.[1] || "";
  const destinationPostalCode = extractPostalCode(destination);

  if (!routePostalCode || !destinationPostalCode || routePostalCode === destinationPostalCode) {
    return destination;
  }

  return normalizeFreeText(destination.replace(destinationPostalCode, routePostalCode));
}

function reconcileRouteCode(routeCode, routeLabel, destination) {
  const normalizedRouteCode = normalizeRouteCode(routeCode || "");
  const postalCode = extractPostalCode(destination);
  const routePrefix = routeLabel.match(/\bR\d+\b/i)?.[0]?.toUpperCase() || "";
  const derivedRouteCode = routePrefix && postalCode ? `${routePrefix}${postalCode}` : "";

  if (!normalizedRouteCode) {
    return derivedRouteCode;
  }

  if (postalCode && !normalizedRouteCode.endsWith(postalCode) && derivedRouteCode) {
    return derivedRouteCode;
  }

  return normalizedRouteCode;
}

function reconcileBarcode(barcode, destination, routeCode) {
  const normalizedBarcode = normalizeBarcode(barcode);
  const postalCode = extractPostalCode(destination);
  const routePostalCode = routeCode.match(/(\d{5})$/)?.[1] || "";

  if (
    normalizedBarcode.length === 5
    && (normalizedBarcode === postalCode || normalizedBarcode === routePostalCode)
  ) {
    return "";
  }

  return normalizedBarcode;
}

function reconcileCommandNumber(commandNumber, barcode) {
  return normalizeCommandNumber(commandNumber) || deriveCommandNumberFromBarcode(barcode);
}

function extractPostalCode(destination) {
  return String(destination).match(/\b\d{5}\b/)?.[0] || "";
}

function getDestinationShortLabel(destination) {
  const cleaned = sanitizeDestination(destination);
  const match = cleaned.match(/\b\d{5}\s+[A-ZÀ-ÿ-]+(?:\s+[A-ZÀ-ÿ-]+){0,2}/u);
  return match ? normalizeFreeText(match[0]) : cleaned;
}

function renderRouteCodeMeta(parcels) {
  const routeCodes = [...new Set(parcels.map((parcel) => parcel.routeCode).filter(Boolean))];
  if (routeCodes.length !== 1) {
    return "";
  }

  return `<span><strong>Route :</strong> ${escapeHtml(formatRouteCodeForDisplay(routeCodes[0]))}</span>`;
}

function normalizeBarcode(value) {
  return value.trim();
}

function parseParcelWeightKg(parcel) {
  const rawWeight = typeof parcel === "string" ? parcel : parcel?.weight || "";
  const match = String(rawWeight).match(/(\d+(?:[.,]\d+)?)/);
  if (!match) {
    return null;
  }

  const parsedWeight = Number.parseFloat(match[1].replace(",", "."));
  return Number.isFinite(parsedWeight) ? parsedWeight : null;
}

function getParcelHandlingFactor(weightKg) {
  if (weightKg === null) {
    return 1.2;
  }

  if (weightKg >= 40) {
    return 4.5;
  }

  if (weightKg >= 25) {
    return 3.5;
  }

  if (weightKg >= 15) {
    return 2;
  }

  if (weightKg >= 5) {
    return 1.25;
  }

  return 1;
}

function getSortingEffortLabel(totalEffort) {
  if (totalEffort >= 18) {
    return "Effort soutenu";
  }

  if (totalEffort >= 8) {
    return "Effort moyen";
  }

  return "Effort leger";
}

function getSortingHandlingAdvice(summary) {
  if (summary.heaviestParcelKg >= 35 || summary.totalWeightKg >= 120) {
    return "Utiliser un chariot ou faire plusieurs depots courts";
  }

  if (summary.heaviestParcelKg >= 25) {
    return "Traiter les colis lourds un par un";
  }

  if (summary.totalWeightKg <= 15 && summary.movedCount > 1 && !summary.unknownWeightCount) {
    return "Collecte groupee possible";
  }

  if (summary.unknownWeightCount) {
    return "Verifier le poids avant de prendre plusieurs colis";
  }

  return "Avancer baque par baque vers la cible";
}

function formatKnownWeightSummary(totalWeightKg, unknownWeightCount = 0) {
  const roundedWeight = Number(totalWeightKg || 0);
  const formattedWeight = roundedWeight > 0
    ? `${roundedWeight.toLocaleString("fr-FR", { minimumFractionDigits: roundedWeight < 100 ? 1 : 0, maximumFractionDigits: 1 })} kg`
    : "";

  if (formattedWeight && unknownWeightCount) {
    return `${formattedWeight} + ${unknownWeightCount} poids inconnus`;
  }

  if (formattedWeight) {
    return formattedWeight;
  }

  if (unknownWeightCount) {
    return `${unknownWeightCount} poids inconnus`;
  }

  return "Poids non renseigne";
}

function normalizeCommandNumber(value) {
  const normalizedValue = normalizeBarcode(String(value || ""));
  return /^\d{5,8}$/.test(normalizedValue) ? normalizedValue : "";
}

function deriveCommandNumberFromBarcode(barcode) {
  const normalizedBarcode = normalizeBarcode(barcode);
  if (normalizeCommandNumber(normalizedBarcode)) {
    return normalizedBarcode;
  }

  if (/^\d{8,11}$/.test(normalizedBarcode)) {
    const truncatedValue = normalizedBarcode.slice(0, -3);
    if (normalizeCommandNumber(truncatedValue)) {
      return truncatedValue;
    }
  }

  return "";
}

function getParcelCommandNumber(parcel) {
  return reconcileCommandNumber(parcel?.commandNumber || "", parcel?.barcode || "");
}

function normalizeDeliveryNote(note) {
  if (!note || !note.id || !note.name || !note.importedAt) {
    return null;
  }

  return {
    id: String(note.id),
    name: normalizeFreeText(String(note.name)),
    size: Number(note.size || 0),
    importedAt: normalizeStoredDate(note.importedAt, new Date().toISOString()),
    analysis: normalizeDeliveryNoteAnalysis(note.analysis),
  };
}

function normalizeDeliveryNoteAnalysis(analysis) {
  if (!analysis || !Array.isArray(analysis.missingEntries)) {
    return null;
  }

  return {
    totalEntries: Number(analysis.totalEntries || 0),
    totalExpectedCount: Number(analysis.totalExpectedCount || 0),
    totalRegisteredCount: Number(analysis.totalRegisteredCount || 0),
    totalMissingCount: Number(analysis.totalMissingCount || 0),
    incomparableParcelsCount: Number(analysis.incomparableParcelsCount || 0),
    parseError: normalizeFreeText(analysis.parseError || ""),
    missingEntries: analysis.missingEntries
      .map((entry) => ({
        commandNumber: normalizeCommandNumber(entry.commandNumber || ""),
        expectedCount: Number(entry.expectedCount || 1),
        registeredCount: Number(entry.registeredCount || 0),
        missingCount: Number(entry.missingCount || 0),
        client: normalizeFreeText(entry.client || ""),
        city: normalizeFreeText(entry.city || ""),
        rawContext: normalizeFreeText(entry.rawContext || ""),
      }))
      .filter((entry) => entry.commandNumber),
    analyzedAt: normalizeStoredDate(analysis.analyzedAt || "", ""),
  };
}

function looksLikePdf(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name || "");
}

function getDeliveryNoteErrorMessage(error, fallbackMessage = "Impossible d'analyser ce PDF.") {
  const rawMessage = normalizeFreeText(String(error?.message || error || ""));
  if (!rawMessage) {
    return fallbackMessage;
  }

  if (/withResolvers/i.test(rawMessage)) {
    return "Le lecteur PDF du telephone n'etait pas compatible. Rechargez la page puis reessayez.";
  }

  if (/pdfjs-script-load-failed|pdfjs-unavailable|failed to fetch dynamically imported module|importing a module script failed|load failed|fetch/i.test(rawMessage)) {
    return "Le lecteur PDF n'a pas pu etre charge. Verifiez la connexion puis rechargez la page.";
  }

  if (/indexeddb/i.test(rawMessage)) {
    return "Le stockage local du PDF est bloque sur ce telephone. Desactivez le mode prive puis reessayez.";
  }

  if (/tesseract-unavailable/i.test(rawMessage)) {
    return "L'OCR du PDF n'est pas disponible sur ce navigateur.";
  }

  if (/delivery-note-empty/i.test(rawMessage)) {
    return "Aucune livraison exploitable n'a ete detectee dans ce PDF.";
  }

  return fallbackMessage;
}

function formatFileSize(size) {
  const bytes = Number(size || 0);
  if (bytes >= 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(1)} Mo`;
  }

  if (bytes >= 1024) {
    return `${Math.round(bytes / 1024)} Ko`;
  }

  return `${bytes} o`;
}

function formatRouteCodeForDisplay(routeCode) {
  const normalized = normalizeRouteCode(routeCode || "");
  const match = normalized.match(/^R(\d+)(\d{5})$/);
  if (!match) {
    return normalized;
  }

  const routePrefix = `R${match[1]}`;
  const postalCode = match[2];
  return `${routePrefix} ${postalCode.slice(0, 2)} ${postalCode.slice(2)}`;
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

function normalizePdfOcrLine(value) {
  return normalizeDeliveryTextLine(
    String(value)
      .replace(/[|]/g, "I")
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'")
      .replace(/[_]+/g, " ")
      .replace(/[^\S\r\n]+/g, " "),
  );
}

function normalizeDeliveryTextLine(value) {
  return normalizeFreeText(
    String(value)
      .replace(/([A-Za-zÀ-ÿ'´])(?=\d{5}\b)/gu, "$1 ")
      .replace(/(\d{5})(?=[A-Za-zÀ-ÿ'´])/gu, "$1 ")
      .replace(/([A-Za-zÀ-ÿ'´])(?=0\d(?:[ .]?\d{2}){4}\b)/gu, "$1 ")
      .replace(/([a-zà-ÿ])([A-ZÀ-Ý])/gu, "$1 $2"),
  );
}

function normalizeDeliveryCommandLine(line) {
  return normalizeDeliveryTextLine(line)
    .replace(/C\s*O\s*M\s*M?\s*A\s*N\s*D\s*E/gi, "COMMANDE")
    .replace(/C\s*D\s*E/gi, "CDE");
}

function parseStructuredDeliveryNoteLines(lines) {
  const entries = [];
  let currentContext = createDeliveryContext();
  let pendingPackageCount = 0;

  lines.forEach((line, index) => {
    if (!line || isDeliveryPageMetaLine(line)) {
      return;
    }

    if (isDeliveryAddressHeaderLine(line)) {
      const blockLines = [];

      for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
        const candidate = lines[cursor];
        if (isDeliveryDetailHeaderLine(candidate) || isDeliveryAddressHeaderLine(candidate) || isDeliveryPageMetaLine(candidate)) {
          break;
        }

        blockLines.push(candidate);
      }

      const nextContext = extractDeliveryContextFromBlock(blockLines);
      if (nextContext.client || nextContext.city) {
        currentContext = nextContext;
      }
      return;
    }

    const packageCount = extractDeliveryPackageCountFromSummaryLine(line);
    if (packageCount) {
      pendingPackageCount = packageCount;
      return;
    }

    const commandNumber = extractDeliveryCommandNumber(line);
    if (!commandNumber) {
      return;
    }

    const barcodeInfo = findNearbyDeliveryBarcode(lines, index, commandNumber);
    const resolvedContext = resolveStructuredDeliveryContext(lines, index, currentContext);
    entries.push({
      commandNumber,
      expectedCount: resolveStructuredDeliveryPackageCount(lines, index, pendingPackageCount),
      client: resolvedContext.client,
      city: resolvedContext.city,
      rawContext: buildDeliveryEntryContext(resolvedContext, line, barcodeInfo?.line || ""),
    });

    pendingPackageCount = 0;
  });

  return entries;
}

function createDeliveryContext() {
  return {
    client: "",
    city: "",
    blockLines: [],
  };
}

function resolveStructuredDeliveryContext(lines, commandIndex, fallbackContext) {
  const nearbyBlockLines = collectNearbyDeliveryContextLines(lines, commandIndex);
  const nearbyContext = nearbyBlockLines.length
    ? extractDeliveryContextFromBlock(nearbyBlockLines)
    : createDeliveryContext();

  return chooseBestDeliveryContext([nearbyContext, fallbackContext]);
}

function collectNearbyDeliveryContextLines(lines, commandIndex) {
  const blockLines = [];
  const startIndex = Math.max(0, commandIndex - 12);

  for (let index = startIndex; index < commandIndex; index += 1) {
    const candidate = lines[index];
    if (!candidate) {
      continue;
    }

    if (isDeliveryPageMetaLine(candidate) || isDeliveryAddressHeaderLine(candidate)) {
      blockLines.length = 0;
      continue;
    }

    if (!looksLikeDeliveryContextLine(candidate)) {
      continue;
    }

    blockLines.push(candidate);
  }

  return blockLines.slice(-6);
}

function looksLikeDeliveryContextLine(line) {
  const normalizedLine = normalizeDeliveryTextLine(line);
  if (!normalizedLine || isDeliveryContextSeparatorLine(normalizedLine)) {
    return false;
  }

  if (!/[A-ZÀ-Ý]/u.test(normalizedLine)) {
    return false;
  }

  return !/^\d+(?:\s+\d+)*$/.test(normalizedLine);
}

function isDeliveryContextSeparatorLine(line) {
  return isDeliveryDetailHeaderLine(line)
    || isDeliveryPageMetaLine(line)
    || isDeliveryAddressHeaderLine(line)
    || Boolean(extractDeliveryCommandNumber(line))
    || Boolean(extractDeliveryBarcodeInfo(line))
    || Boolean(extractDeliveryPackageCountFromSummaryLine(line))
    || /\bPOIDS\b|\bKG\b/i.test(line);
}

function chooseBestDeliveryContext(candidates) {
  const normalizedCandidates = candidates
    .map((candidate) => normalizeDeliveryContextCandidate(candidate))
    .filter((candidate) => candidate.client || candidate.city || candidate.blockLines.length);

  if (!normalizedCandidates.length) {
    return createDeliveryContext();
  }

  return normalizedCandidates.sort((left, right) => scoreDeliveryContext(right) - scoreDeliveryContext(left))[0];
}

function normalizeDeliveryContextCandidate(candidate) {
  return {
    client: sanitizeDeliveryClientLabel(candidate?.client || ""),
    city: sanitizeDeliveryCityLabel(candidate?.city || ""),
    blockLines: Array.isArray(candidate?.blockLines) ? candidate.blockLines.filter(Boolean) : [],
  };
}

function scoreDeliveryContext(candidate) {
  let score = candidate.blockLines.length;

  if (candidate.client) {
    score += 18 + getDeliveryTextQualityScore(candidate.client, "client");
  }

  if (candidate.city) {
    score += 18 + getDeliveryTextQualityScore(candidate.city, "city");
  }

  if (candidate.client && candidate.city && candidate.client !== candidate.city) {
    score += 6;
  }

  return score;
}

function getDeliveryTextQualityScore(value, kind) {
  const normalizedValue = normalizeFreeText(value);
  if (!normalizedValue) {
    return -100;
  }

  let score = 0;
  const tokens = normalizedValue.split(" ").filter(Boolean);

  if (tokens.length >= 2) {
    score += 6;
  }

  if (kind === "client" && /\b(?:SARL|SAS|SASU|EURL|SA|SNC)\b/i.test(normalizedValue)) {
    score += 12;
  }

  if (kind === "city" && /[-']/u.test(normalizedValue)) {
    score += 3;
  }

  if (looksSuspiciousDeliveryText(normalizedValue, kind)) {
    score -= 30;
  }

  return score;
}

function looksSuspiciousDeliveryText(value, kind = "generic") {
  const normalizedValue = normalizeFreeText(value);
  if (!normalizedValue) {
    return true;
  }

  if (/\d/.test(normalizedValue)) {
    return true;
  }

  const tokens = normalizedValue.split(" ").filter(Boolean);
  if (!tokens.length) {
    return true;
  }

  if (hasRepeatedDeliveryTokens(tokens)) {
    return true;
  }

  if (
    kind !== "client"
    && tokens.length > 1
    && tokens.every((token) => normalizeDeliveryTokenKey(token).length <= 3)
  ) {
    return true;
  }

  return false;
}

function hasRepeatedDeliveryTokens(tokens) {
  return tokens.some((token, index) => {
    if (index === 0) {
      return false;
    }

    return normalizeDeliveryTokenKey(token) === normalizeDeliveryTokenKey(tokens[index - 1]);
  });
}

function normalizeDeliveryTokenKey(value) {
  return String(value).replace(/[^A-ZÀ-Ý]/giu, "").toUpperCase();
}

function isDeliveryAddressHeaderLine(line) {
  return /CODE\s*POSTAL/i.test(line) && /VILLE/i.test(line) && /CLIENT/i.test(line);
}

function isDeliveryDetailHeaderLine(line) {
  return /N[º°O]?\s*B/i.test(line) && /N[º°O]?\s*COLIS/i.test(line);
}

function isDeliveryPageMetaLine(line) {
  return [
    /GAVIOTA FRANCE/i,
    /BON DE LIVRAISON/i,
    /N[º°O]?\s*DISTRIBUTION/i,
    /^DU\s+\d{2}\/\d{2}\/\d{4}/i,
    /^PAGE\s+\d+\s*\/\s*\d+/i,
    /^TOTAL\s+COLIS/i,
  ].some((regex) => regex.test(line));
}

function extractDeliveryPackageCountFromSummaryLine(line) {
  const match = normalizeDeliveryTextLine(line).match(/^\d{6}\s+\d{2}\/\d{2}\/\d{4}\s+([1-9]\d?)[,.]0{1,2}\b/);
  return match ? Number.parseInt(match[1], 10) : 0;
}

function extractDeliveryCommandNumber(line) {
  const normalizedLine = normalizeDeliveryCommandLine(line);
  const match = normalizedLine.match(/COMMANDE[^\d]{0,20}(\d{5,8})\b/i);
  return match ? normalizeBarcode(match[1]) : "";
}

function extractDeliveryBarcodeInfo(line) {
  const match = normalizeDeliveryTextLine(line).match(/\*?(\d{5,8})(\d{3})\*?/);
  if (!match) {
    return null;
  }

  return {
    commandNumber: normalizeBarcode(match[1]),
    packageCount: Number.parseInt(match[2], 10),
    line: normalizeDeliveryTextLine(line),
  };
}

function findNearbyDeliveryBarcode(lines, startIndex, expectedCommandNumber = "") {
  for (let index = startIndex + 1; index <= Math.min(lines.length - 1, startIndex + 3); index += 1) {
    const barcodeInfo = extractDeliveryBarcodeInfo(lines[index]);
    if (!barcodeInfo) {
      continue;
    }

    if (expectedCommandNumber && barcodeInfo.commandNumber !== expectedCommandNumber) {
      continue;
    }

    return barcodeInfo;
  }

  return null;
}

function resolveStructuredDeliveryPackageCount(lines, commandIndex, pendingPackageCount) {
  if (pendingPackageCount) {
    return pendingPackageCount;
  }

  const startIndex = Math.max(0, commandIndex - 6);
  const endIndex = Math.min(lines.length, commandIndex + 4);
  return extractDeliveryPackageCount(lines.slice(startIndex, endIndex), commandIndex - startIndex);
}

function buildDeliveryEntryContext(context, commandLine, barcodeLine) {
  return [context.client, context.city, ...context.blockLines.slice(0, 2), commandLine, barcodeLine]
    .filter(Boolean)
    .join(" | ");
}

function resolveDeliveryCommandNumber(initialValue, contextLines) {
  const frequencies = new Map();
  const candidates = contextLines.flatMap((line, index) => {
    const matches = [...line.matchAll(/\b\d{5,8}\b/g)].map((match) => ({
      value: normalizeBarcode(match[0]),
      line,
      index,
    }));

    matches.forEach((candidate) => {
      frequencies.set(candidate.value, (frequencies.get(candidate.value) || 0) + 1);
    });

    return matches;
  });

  const bestCandidate = candidates
    .filter((candidate) => !looksLikeDeliveryDate(candidate.value))
    .sort((left, right) => getDeliveryCommandCandidateScore(right, initialValue, frequencies) - getDeliveryCommandCandidateScore(left, initialValue, frequencies))[0];

  return bestCandidate?.value || initialValue;
}

function getDeliveryCommandCandidateScore(candidate, initialValue, frequencies) {
  let score = 0;
  const line = normalizeDeliveryCommandLine(candidate.line || "");

  if (candidate.value === initialValue) {
    score += 20;
  }

  if (/COMMANDE/i.test(line)) {
    score += 40;
  }

  if (/\bCDE\b|CODE/i.test(line)) {
    score += 30;
  }

  if (candidate.value.length >= 6 && candidate.value.length <= 7) {
    score += 10;
  }

  if (candidate.value.startsWith("0")) {
    score += 4;
  }

  score += (frequencies.get(candidate.value) || 0) * 8;
  score -= candidate.index * 2;

  return score;
}

function extractDeliveryClient(lines) {
  return extractDeliveryClientFromBlock(lines);
}

function sliceUppercaseClientTokens(value) {
  const tokens = normalizeDeliveryTextLine(value).split(" ");
  const clientTokens = [];

  for (const token of tokens) {
    if (/^\d{5}$/.test(token)) {
      break;
    }

    if (/^\d+([,./-]\S*)?$/.test(token) || token.includes("/")) {
      break;
    }

    if (/\d/.test(token)) {
      break;
    }

    if (isDeliveryAddressToken(token)) {
      break;
    }

    if (/[a-zà-ÿ]/u.test(token) && !looksLikeAllowedLowercaseClientToken(token)) {
      break;
    }

    clientTokens.push(token);
  }

  const clientName = normalizeDeliveryClientName(clientTokens.join(" "));
  return /[A-Z]/.test(clientName) ? clientName : "";
}

function extractDeliveryCity(lines) {
  return extractDeliveryCityFromBlock(lines);
}

function extractDeliveryContextFromBlock(lines) {
  const blockLines = lines
    .filter(Boolean)
    .map((line) => normalizeDeliveryTextLine(line));

  return {
    client: extractDeliveryClientFromBlock(blockLines),
    city: extractDeliveryCityFromBlock(blockLines),
    blockLines,
  };
}

function extractDeliveryClientFromBlock(lines) {
  const candidates = [];

  lines.forEach((line, index) => {
    const normalizedLine = normalizeDeliveryTextLine(line.replace(/[©@]/g, " "));
    const payload = normalizedLine.replace(/^(?:\d+\s+)?\d{5,6}\s+/, "");

    if (payload === normalizedLine && /\b\d{5}\b/.test(normalizedLine)) {
      return;
    }

    const clientName = sanitizeDeliveryClientLabel(sliceUppercaseClientTokens(payload));
    if (clientName) {
      candidates.push({
        value: clientName,
        score: getDeliveryTextQualityScore(clientName, "client") - index,
      });
    }
  });

  return candidates.sort((left, right) => right.score - left.score)[0]?.value || "";
}

function extractDeliveryCityFromBlock(lines) {
  const normalizedLines = lines.map((line) => normalizeDeliveryTextLine(line));
  const candidates = [];

  for (let index = 0; index < normalizedLines.length; index += 1) {
    const line = normalizedLines[index];
    const postalMatches = [...line.matchAll(/\b\d{5}\b/g)];
    const postalMatch = postalMatches[postalMatches.length - 1];
    if (!postalMatch) {
      continue;
    }

    let city = collectDeliveryCityTokens(line.slice(postalMatch.index + postalMatch[0].length));

    if (index > 0) {
      const previousLine = normalizedLines[index - 1];
      if (!/\b\d{5}\b/.test(previousLine) || looksLikeFrenchPhoneNumber(previousLine)) {
        const previousFragment = collectDeliveryTrailingUppercaseFragment(previousLine);
        if (previousFragment && (previousFragment.length <= 10 || !city || city.length <= 10)) {
          city = joinDeliveryCityFragments(previousFragment, city);
        }
      }
    }

    for (let cursor = index + 1; cursor < normalizedLines.length; cursor += 1) {
      const nextLine = normalizedLines[cursor];
      if (isDeliveryDetailHeaderLine(nextLine) || isDeliveryAddressHeaderLine(nextLine) || isDeliveryPageMetaLine(nextLine)) {
        break;
      }

      if (/\b\d{5}\b/.test(nextLine) || looksLikeFrenchPhoneNumber(nextLine)) {
        break;
      }

      if (/[a-zà-ÿ]/u.test(nextLine)) {
        break;
      }

      const continuation = collectDeliveryCityTokens(nextLine);
      if (!continuation) {
        break;
      }

      city = joinDeliveryCityFragments(city, continuation);
    }

    if (city) {
      const sanitizedCity = sanitizeDeliveryCityLabel(city);
      if (sanitizedCity) {
        candidates.push({
          value: sanitizedCity,
          score: getDeliveryTextQualityScore(sanitizedCity, "city") - index,
        });
      }
    }
  }

  return candidates.sort((left, right) => right.score - left.score)[0]?.value || "";
}

function normalizeDeliveryClientName(value) {
  return normalizeDeliveryTextLine(value)
    .replace(/([A-ZÀ-Ý]{3,})(SARL|SAS|SASU|EURL)\b/gu, "$1 $2");
}

function sanitizeDeliveryClientLabel(value) {
  const cleaned = collapseRepeatedDeliveryTokens(
    normalizeDeliveryClientName(value)
      .split(" ")
      .filter((token) => token && !/\d/.test(token))
      .join(" "),
  );

  return looksSuspiciousDeliveryText(cleaned, "client") ? "" : cleaned;
}

function sanitizeDeliveryCityLabel(value) {
  const cleaned = collapseRepeatedDeliveryTokens(
    normalizeDeliveryTextLine(value)
      .split(" ")
      .filter((token) => token && !/\d/.test(token))
      .join(" "),
  );

  return looksSuspiciousDeliveryText(cleaned, "city") ? "" : cleaned;
}

function collapseRepeatedDeliveryTokens(value) {
  const tokens = normalizeFreeText(value).split(" ").filter(Boolean);
  const collapsedTokens = [];

  tokens.forEach((token) => {
    if (!collapsedTokens.length || normalizeDeliveryTokenKey(collapsedTokens[collapsedTokens.length - 1]) !== normalizeDeliveryTokenKey(token)) {
      collapsedTokens.push(token);
    }
  });

  return normalizeFreeText(collapsedTokens.join(" "));
}

function looksLikeAllowedLowercaseClientToken(token) {
  return /^(sarl|sas|sasu|eurl|sa|snc|mr|mme|mlle)$/i.test(token);
}

function isDeliveryAddressToken(token) {
  return /^(?:ZA|ZI|ZAC|ZD|ZONE|PARC|RUE|AVENUE|AVE|CHEMIN|IMPASSE|ROUTE|BD|BOULEVARD|ALLEE|ALLEES|RESIDENCE|RÉSIDENCE|BAT|BÂT|QUAI|LIEU-DIT|LIEUDIT)$/i.test(token);
}

function collectDeliveryCityTokens(value) {
  const tokens = normalizeDeliveryTextLine(value)
    .replace(/\b0\d(?:[ .]?\d{2}){4}\b/g, " ")
    .split(" ")
    .filter(Boolean);
  const cityTokens = [];

  for (const token of tokens) {
    if (/^\d+$/.test(token) || /\d/.test(token) || looksLikeFrenchPhoneNumber(token)) {
      break;
    }

    if (!/[A-ZÀ-Ý]/u.test(token)) {
      break;
    }

    cityTokens.push(token);
  }

  return normalizeFreeText(cityTokens.join(" "));
}

function collectDeliveryTrailingUppercaseFragment(value) {
  const cleaned = normalizeDeliveryTextLine(value)
    .replace(/\b0\d(?:[ .]?\d{2}){4}\b/g, " ")
    .trim();
  const match = cleaned.match(/([A-ZÀ-Ý'´&-]+(?:\s+[A-ZÀ-Ý'´&-]+)*)$/u);
  return match ? normalizeFreeText(match[1]) : "";
}

function joinDeliveryCityFragments(left, right) {
  if (!left) {
    return sanitizeDeliveryCityLabel(right);
  }

  if (!right) {
    return sanitizeDeliveryCityLabel(left);
  }

  if (right.length === 1 || /-$/.test(left) || /-[A-ZÀ-Ý]+$/u.test(left)) {
    return sanitizeDeliveryCityLabel(`${left}${right}`);
  }

  return sanitizeDeliveryCityLabel(`${left} ${right}`);
}

function looksLikeFrenchPhoneNumber(value) {
  return /\b0\d(?:[ .]?\d{2}){4}\b/.test(value);
}

function extractDeliveryPackageCount(lines, commandIndex) {
  const candidates = [];

  lines.forEach((line, index) => {
    const normalizedLine = normalizeDeliveryTextLine(line.replace(/\bO,00\b/gi, "0,00"));
    const summaryCount = extractDeliveryPackageCountFromSummaryLine(normalizedLine);
    if (summaryCount) {
      candidates.push({
        count: summaryCount,
        line: normalizedLine,
        index,
        source: "summary",
      });
    }

    const decimalMatches = [...normalizedLine.matchAll(/\b([1-9]\d?)[,.]0{1,2}\b/g)];
    decimalMatches.forEach((match) => {
      candidates.push({
        count: Number.parseInt(match[1], 10),
        line: normalizedLine,
        index,
        source: "decimal",
      });
    });

    if (/\d{2}\/\d{2}\/\d{4}/.test(normalizedLine)) {
      const compactMatches = [...normalizedLine.matchAll(/\b([1-9]\d?)00\b/g)];
      compactMatches.forEach((match) => {
        candidates.push({
          count: Number.parseInt(match[1], 10),
          line: normalizedLine,
          index,
          source: "compact",
        });
      });
    }
  });

  const bestCandidate = candidates
    .filter((candidate) => candidate.count > 0 && candidate.count < 20)
    .sort((left, right) => getDeliveryPackageCountScore(right, commandIndex) - getDeliveryPackageCountScore(left, commandIndex))[0];

  return bestCandidate?.count || 1;
}

function getDeliveryPackageCountScore(candidate, commandIndex) {
  let score = 0;
  const line = candidate.line || "";

  if (/COLIS/i.test(line)) {
    score += 20;
  }

  if (/\d{2}\/\d{2}\/\d{4}/.test(line)) {
    score += 16;
  }

  if (/POIDS|KG|\bX\b/i.test(line)) {
    score -= 18;
  }

  if (candidate.source === "decimal") {
    score += 12;
  }

  if (candidate.source === "summary") {
    score += 28;
  }

  score -= Math.abs(candidate.index - commandIndex) * 4;
  return score;
}

function dedupeDeliveryEntries(entries) {
  const bestEntries = new Map();

  entries.forEach((entry, index) => {
    if (!entry.commandNumber) {
      return;
    }

    const current = bestEntries.get(entry.commandNumber);
    const candidateScore = scoreDeliveryEntryCandidate(entry);

    if (!current || candidateScore > current.score) {
      bestEntries.set(entry.commandNumber, {
        entry,
        score: candidateScore,
        orderIndex: current?.orderIndex ?? index,
      });
    }
  });

  return [...bestEntries.values()]
    .sort((left, right) => left.orderIndex - right.orderIndex)
    .map((item) => item.entry);
}

function scoreDeliveryEntryCandidate(entry) {
  return scoreDeliveryContext({
    client: entry.client || "",
    city: entry.city || "",
    blockLines: [],
  }) + (Number(entry.expectedCount || 0) > 0 && Number(entry.expectedCount || 0) < 10 ? 4 : 0);
}

function buildRegisteredCommandSet() {
  return new Set(
    state.parcels
      .map((parcel) => getParcelCommandNumber(parcel))
      .filter(Boolean),
  );
}

function buildRegisteredCommandCounts() {
  const groupedCounts = state.parcels.reduce((map, parcel) => {
    const commandNumber = getParcelCommandNumber(parcel);
    if (!commandNumber) {
      return map;
    }

    const bucket = map.get(commandNumber) || {
      packageIndexes: new Set(),
      unlabeledCount: 0,
    };
    const packageIndex = normalizeFreeText(parcel.packageIndex || "");

    if (packageIndex) {
      bucket.packageIndexes.add(packageIndex);
    } else {
      bucket.unlabeledCount += 1;
    }

    map.set(commandNumber, bucket);
    return map;
  }, new Map());

  return new Map(
    [...groupedCounts.entries()].map(([commandNumber, bucket]) => [
      commandNumber,
      bucket.packageIndexes.size + bucket.unlabeledCount,
    ]),
  );
}

function buildRegisteredCommandInfo() {
  const groupedInfo = state.parcels.reduce((map, parcel) => {
    const commandNumber = getParcelCommandNumber(parcel);
    if (!commandNumber) {
      return map;
    }

    const nextBucket = map.get(commandNumber) || {
      clients: new Map(),
      cities: new Map(),
    };
    const client = normalizeFreeText(parcel.client || "");
    const city = getParcelCityLabel(parcel);

    if (client) {
      nextBucket.clients.set(client, (nextBucket.clients.get(client) || 0) + 1);
    }

    if (city) {
      nextBucket.cities.set(city, (nextBucket.cities.get(city) || 0) + 1);
    }

    map.set(commandNumber, nextBucket);
    return map;
  }, new Map());

  return new Map(
    [...groupedInfo.entries()].map(([commandNumber, bucket]) => [
      commandNumber,
      {
        client: getMostCommonDeliveryValue(bucket.clients),
        city: getMostCommonDeliveryValue(bucket.cities),
      },
    ]),
  );
}

function countIncomparableParcels() {
  return state.parcels.filter((parcel) => !getParcelCommandNumber(parcel)).length;
}

function getMostCommonDeliveryValue(counts) {
  return [...counts.entries()]
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0], "fr", { sensitivity: "base" }))[0]?.[0] || "";
}

function getParcelCityLabel(parcel) {
  const destination = sanitizeDestination(parcel.destination || "");
  const match = destination.match(/\b\d{5}\s+(.+)$/);
  return normalizeFreeText(match ? match[1] : destination);
}

function looksLikeDeliveryDate(value) {
  return /^(?:19|20)\d{6}$/.test(value) || /^(?:0[1-9]|[12]\d|3[01])(?:0[1-9]|1[0-2])\d{2}$/.test(value);
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
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

function formatDate(value) {
  const normalizedDate = normalizeStoredDate(value, "");
  if (!normalizedDate) {
    return "Date inconnue";
  }

  try {
    return new Intl.DateTimeFormat("fr-FR", {
      dateStyle: "short",
      timeStyle: "short",
    }).format(new Date(normalizedDate));
  } catch (error) {
    return new Date(normalizedDate).toLocaleString("fr-FR");
  }
}

function pluralize(count, singular, plural) {
  return count > 1 ? plural : singular;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

function normalizeStoredDate(value, fallback = "") {
  if (!value) {
    return fallback;
  }

  const parsedDate = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? fallback : parsedDate.toISOString();
}
