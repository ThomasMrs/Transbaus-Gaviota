const STORAGE_KEY = "transbaus-gaviota-state-v1";
const COLLAPSE_STORAGE_KEY = "le-baus-du-tri-collapse-v1";
const ACCESS_STORAGE_KEY = "transbaus-gaviota-access-v1";
const ACCESS_RATE_LIMIT_STORAGE_KEY = "transbaus-gaviota-access-rate-v1";
const ACCESS_PASSWORD = "2005";
const ACCESS_FAILED_ATTEMPTS_LIMIT = 3;
const ACCESS_LOCK_DURATION_MS = 10_000;
const PDF_DB_NAME = "le-baus-du-tri-documents-v1";
const PDF_STORE_NAME = "delivery-notes";
const PDFJS_VERSION = "5.6.205";
const PDFJS_MODULE_URL = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.mjs`;
const PDFJS_WORKER_URL = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.mjs`;
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
  ui.destinationSummary = document.querySelector("#destinationSummary");
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
      createdAt: baque.createdAt || new Date().toISOString(),
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
        createdAt: parcel.createdAt || new Date().toISOString(),
        updatedAt: parcel.updatedAt || parcel.createdAt || new Date().toISOString(),
      }))
      .filter((parcel) => parcel.routeCode || parcel.commandNumber || parcel.barcode || parcel.destination);
    const deliveryNotes = Array.isArray(parsed.deliveryNotes)
      ? parsed.deliveryNotes
        .map((note) => normalizeDeliveryNote(note))
        .filter(Boolean)
      : [];

    return {
      baques: baques.length ? baques : createDefaultState().baques,
      parcels,
      deliveryNotes,
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
      createdAt: new Date().toISOString(),
    })),
    parcels: [],
    deliveryNotes: [],
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
  renderDestinationSummary();
  renderBaques();
  renderSearchResults();
  renderDeliveryNotes();
  applyCollapseStateToDom();
}

function renderHeroStats() {
  const totalBaques = state.baques.length;
  const totalParcels = state.parcels.length;
  const totalDestinations = new Set(
    state.parcels
      .map((parcel) => getParcelDestinationKey(parcel))
      .filter(Boolean),
  ).size;

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

  ui.parcelBaqueSelect.innerHTML = state.baques
    .map(
      (baque) => `
        <option value="${escapeHtml(baque.id)}">
          ${escapeHtml(baque.name)} - ${escapeHtml(baque.location)}
        </option>
      `,
    )
    .join("");

  ui.parcelBaqueSelect.value = state.baques.some((baque) => baque.id === previousValue)
    ? previousValue
    : state.baques[0]?.id || "";
}

function renderDestinationSummary() {
  if (!state.parcels.length) {
    ui.destinationSummary.innerHTML = `
      <article class="empty-card">
        <p class="empty-state">Aucun colis pour le moment. La vue par destination apparaitra ici.</p>
      </article>
    `;
    return;
  }

  const grouped = state.parcels.reduce((map, parcel) => {
    const key = getParcelDestinationKey(parcel);
    if (!key) {
      return map;
    }
    if (!map.has(key)) {
      map.set(key, []);
    }
    map.get(key).push(parcel);
    return map;
  }, new Map());

  ui.destinationSummary.innerHTML = [...grouped.entries()]
    .sort(([a], [b]) => a.localeCompare(b, "fr", { numeric: true }))
    .map(([destination, parcels]) => {
      const distribution = parcels.reduce((acc, parcel) => {
        const baqueName = getBaqueById(parcel.currentBaqueId)?.name || "Baque supprimee";
        acc[baqueName] = (acc[baqueName] || 0) + 1;
        return acc;
      }, {});

      const chips = Object.entries(distribution)
        .map(
          ([baqueName, count]) => `
            <span class="distribution-chip">${escapeHtml(baqueName)} : ${escapeHtml(String(count))}</span>
          `,
        )
        .join("");

      return `
        <article class="destination-card">
          <h3>${escapeHtml(destination)}</h3>
          <div class="destination-count">${escapeHtml(String(parcels.length))}</div>
          <div class="destination-card__meta">
            <span>${escapeHtml(pluralize(parcels.length, "colis", "colis"))}</span>
            ${renderRouteCodeMeta(parcels)}
          </div>
          <div class="distribution-list">${chips}</div>
        </article>
      `;
    })
    .join("");
}

function renderBaques() {
  ui.baquesGrid.innerHTML = state.baques
    .map((baque) => {
      const parcels = getParcelsForBaque(baque.id);

      return `
        <article class="baque-card" data-baque-id="${escapeHtml(baque.id)}">
          <div class="baque-card__top">
            <div class="baque-card__meta">
              <span class="count-pill">${escapeHtml(String(parcels.length))} ${escapeHtml(pluralize(parcels.length, "colis", "colis"))}</span>
              <button class="btn btn--danger" type="button" data-action="delete-baque" data-baque-id="${escapeHtml(baque.id)}">
                Supprimer la baque
              </button>
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
          </div>

          <div class="parcel-list">
            ${parcels.length ? parcels.map((parcel) => parcelTemplate(parcel)).join("") : emptyBaqueTemplate()}
          </div>
        </article>
      `;
    })
    .join("");
}

function parcelTemplate(parcel) {
  const options = state.baques
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
    createdAt: new Date().toISOString(),
  });

  saveState();
  render();
  ui.baqueForm.reset();
  showToast(`La baque "${name}" a ete ajoutee.`);
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
      ui.deliveryNoteStatus.textContent = `PDF importe : ${deliveryNote.name}. Analyse impossible pour le moment.`;
      showToast(`PDF ${deliveryNote.name} importe, mais l'analyse a echoue.`, "danger");
    }
  } catch (error) {
    ui.deliveryNoteStatus.textContent = "Impossible d'importer ce PDF.";
    showToast("Impossible d'importer ce PDF.", "danger");
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
      ui.deliveryNoteStatus.textContent = "Impossible d'analyser ce PDF.";
      showToast("Impossible d'analyser ce PDF.", "danger");
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
    ? "Placez l'etiquette entiere dans le cadre, bien droite et nette, puis prenez la photo."
    : "Placez le code-barres au centre du cadre, evitez les reflets, puis prenez la photo.";
  ui.takeCaptureBtn.textContent = "Prendre la photo";
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
  if (captureSession.stream) {
    captureSession.stream.getTracks().forEach((track) => track.stop());
    captureSession.stream = null;
  }

  if (ui.captureVideo) {
    ui.captureVideo.pause();
    ui.captureVideo.srcObject = null;
  }
}

async function handleCapturePhoto() {
  if (captureSession.busy || !ui.captureVideo.videoWidth || !ui.captureVideo.videoHeight) {
    return;
  }

  captureSession.busy = true;
  ui.takeCaptureBtn.disabled = true;
  ui.takeCaptureBtn.textContent = "Preparation...";
  ui.captureStatus.textContent = "Photo prise. Preparation de l'image...";

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
    ui.takeCaptureBtn.textContent = "Prendre la photo";
  }
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
  saveState();
  render();
  showToast("Baque supprimee.");
}

function deleteParcel(parcelId) {
  const parcel = state.parcels.find((item) => item.id === parcelId);
  if (!parcel) {
    return;
  }

  if (!window.confirm(`Supprimer le colis ${getParcelIdentifier(parcel)} ?`)) {
    return;
  }

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

  parcel.currentBaqueId = nextBaqueId;
  parcel.updatedAt = new Date().toISOString();
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

function getOriginLabel(parcel) {
  return getBaqueById(parcel.originBaqueId)?.name || parcel.originBaqueLabel || "Baque supprimee";
}

function createId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }

  return `id-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
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

function getParcelDestinationDisplay(parcel) {
  return parcel.destination || formatRouteCodeForDisplay(parcel.routeCode) || "Sans destination";
}

function getParcelDestinationKey(parcel) {
  return getParcelDestinationDisplay(parcel).toUpperCase();
}

function getParcelIdentifier(parcel) {
  return getParcelCommandNumber(parcel) || parcel.barcode || formatRouteCodeForDisplay(parcel.routeCode) || "Sans code";
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
    importedAt: note.importedAt,
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
    analyzedAt: analysis.analyzedAt || "",
  };
}

function looksLikePdf(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name || "");
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
  if (!pdfjsLibPromise) {
    pdfjsLibPromise = import(PDFJS_MODULE_URL)
      .then((module) => {
        module.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
        return module;
      })
      .catch((error) => {
        pdfjsLibPromise = null;
        throw error;
      });
  }

  return pdfjsLibPromise;
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
  return new Intl.DateTimeFormat("fr-FR", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(new Date(value));
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
