const STORAGE_KEY = "transbaus-gaviota-state-v1";
const DEFAULT_BAQUES = [
  { name: "Baque 1", location: "Zone A" },
  { name: "Baque 2", location: "Zone B" },
  { name: "Baque 3", location: "Zone C" },
  { name: "Baque 4", location: "Zone D" },
];

const state = loadState();
const ui = {};
const scanner = {
  instance: null,
  active: false,
  handled: false,
};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  render();
});

function cacheElements() {
  ui.heroStats = document.querySelector("#heroStats");
  ui.parcelForm = document.querySelector("#parcelForm");
  ui.parcelBaqueSelect = document.querySelector("#parcelBaqueSelect");
  ui.routeCodeInput = document.querySelector("#routeCodeInput");
  ui.destinationInput = document.querySelector("#destinationInput");
  ui.barcodeInput = document.querySelector("#barcodeInput");
  ui.openScannerBtn = document.querySelector("#openScannerBtn");
  ui.baqueForm = document.querySelector("#baqueForm");
  ui.baqueNameInput = document.querySelector("#baqueNameInput");
  ui.baqueLocationInput = document.querySelector("#baqueLocationInput");
  ui.searchInput = document.querySelector("#searchInput");
  ui.searchResults = document.querySelector("#searchResults");
  ui.destinationSummary = document.querySelector("#destinationSummary");
  ui.baquesGrid = document.querySelector("#baquesGrid");
  ui.scannerModal = document.querySelector("#scannerModal");
  ui.reader = document.querySelector("#reader");
  ui.scannerStatus = document.querySelector("#scannerStatus");
  ui.closeScannerBtn = document.querySelector("#closeScannerBtn");
  ui.toastZone = document.querySelector("#toastZone");
}

function bindEvents() {
  ui.parcelForm.addEventListener("submit", handleParcelSubmit);
  ui.baqueForm.addEventListener("submit", handleBaqueSubmit);
  ui.searchInput.addEventListener("input", renderSearchResults);
  ui.openScannerBtn.addEventListener("click", openScanner);
  ui.closeScannerBtn.addEventListener("click", closeScanner);
  ui.scannerModal.addEventListener("click", handleModalClick);
  ui.baquesGrid.addEventListener("click", handleBaqueGridClick);
  ui.baquesGrid.addEventListener("change", handleBaqueGridChange);
  window.addEventListener("beforeunload", () => {
    void stopScanner();
  });
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
      .map((parcel) => ({
        id: parcel.id || createId(),
        barcode: String(parcel.barcode || "").trim(),
        routeCode: String(parcel.routeCode || "").trim().toUpperCase(),
        destination: String(parcel.destination || "").trim(),
        currentBaqueId: parcel.currentBaqueId,
        originBaqueId: parcel.originBaqueId || parcel.currentBaqueId,
        originBaqueLabel: String(parcel.originBaqueLabel || ""),
        createdAt: parcel.createdAt || new Date().toISOString(),
        updatedAt: parcel.updatedAt || parcel.createdAt || new Date().toISOString(),
      }))
      .filter((parcel) => parcel.barcode && parcel.destination);

    return {
      baques: baques.length ? baques : createDefaultState().baques,
      parcels,
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
  };
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
}

function renderHeroStats() {
  const totalBaques = state.baques.length;
  const totalParcels = state.parcels.length;
  const totalDestinations = new Set(state.parcels.map((parcel) => parcel.destination)).size;

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
    if (!map.has(parcel.destination)) {
      map.set(parcel.destination, []);
    }
    map.get(parcel.destination).push(parcel);
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
          <h3>Destination ${escapeHtml(destination)}</h3>
          <div class="destination-count">${escapeHtml(String(parcels.length))}</div>
          <div class="destination-card__meta">
            <span>${escapeHtml(pluralize(parcels.length, "colis", "colis"))}</span>
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

  return `
    <article class="parcel-item" data-parcel-id="${escapeHtml(parcel.id)}">
      <div class="parcel-item__top">
        <div>
          <p class="parcel-code">${escapeHtml(parcel.barcode)}</p>
          <p class="parcel-meta">
            Destination <strong>${escapeHtml(parcel.destination)}</strong><br>
            ${parcel.routeCode ? `Numero destination : ${escapeHtml(parcel.routeCode)}<br>` : ""}
            Origine : ${escapeHtml(getOriginLabel(parcel))}<br>
            Derniere mise a jour : ${escapeHtml(formatDate(parcel.updatedAt || parcel.createdAt))}
          </p>
        </div>
        <span class="tag">Destination ${escapeHtml(parcel.destination)}</span>
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
      parcel.barcode,
      parcel.routeCode || "",
      parcel.destination,
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

      return `
        <article class="search-card">
          <h3>${escapeHtml(parcel.barcode)}</h3>
          <div class="search-card__meta">
            ${parcel.routeCode ? `<span><strong>Numero destination :</strong> ${escapeHtml(parcel.routeCode)}</span>` : ""}
            <span><strong>Destination :</strong> ${escapeHtml(parcel.destination)}</span>
            <span><strong>Baque actuelle :</strong> ${escapeHtml(baque?.name || "Baque supprimee")}</span>
            <span><strong>Emplacement :</strong> ${escapeHtml(baque?.location || "Inconnu")}</span>
            <span><strong>Origine :</strong> ${escapeHtml(getOriginLabel(parcel))}</span>
          </div>
        </article>
      `;
    })
    .join("");
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
}

function upsertParcel(scannedBarcode = "") {
  const baqueId = ui.parcelBaqueSelect.value;
  const routeCode = normalizeRouteCode(ui.routeCodeInput.value);
  const destination = normalizeDestination(ui.destinationInput.value);
  const barcode = normalizeBarcode(scannedBarcode || ui.barcodeInput.value);

  if (!baqueId || !destination || !barcode) {
    showToast("Choisissez une baque, une destination et un code-barres.", "danger");
    return false;
  }

  const baque = getBaqueById(baqueId);
  if (!baque) {
    showToast("La baque selectionnee est introuvable.", "danger");
    return false;
  }

  const now = new Date().toISOString();
  const existing = state.parcels.find((parcel) => parcel.barcode === barcode);

  if (existing) {
    const moved = existing.currentBaqueId !== baqueId;
    existing.routeCode = routeCode;
    existing.destination = destination;
    existing.currentBaqueId = baqueId;
    existing.updatedAt = now;

    saveState();
    render();
    clearParcelForm();
    showToast(
      moved
        ? `Colis ${barcode} deplace vers ${baque.name}.`
        : `Colis ${barcode} mis a jour.`,
    );
    return true;
  }

  state.parcels.unshift({
    id: createId(),
    barcode,
    routeCode,
    destination,
    currentBaqueId: baqueId,
    originBaqueId: baqueId,
    originBaqueLabel: baque.name,
    createdAt: now,
    updatedAt: now,
  });

  saveState();
  render();
  clearParcelForm();
  showToast(`Colis ${barcode} ajoute dans ${baque.name}.`);
  return true;
}

function clearParcelForm() {
  ui.routeCodeInput.value = "";
  ui.destinationInput.value = "";
  ui.barcodeInput.value = "";
  ui.destinationInput.focus();
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

  if (!window.confirm(`Supprimer le colis ${parcel.barcode} ?`)) {
    return;
  }

  state.parcels = state.parcels.filter((item) => item.id !== parcelId);
  saveState();
  render();
  showToast(`Colis ${parcel.barcode} supprime.`);
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
  showToast(`Colis ${parcel.barcode} deplace vers ${nextBaque.name}.`);
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
      const destinationCompare = a.destination.localeCompare(b.destination, "fr", { numeric: true });
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
  return value.trim().replace(/\s+/g, " ");
}

function normalizeRouteCode(value) {
  return value.trim().toUpperCase();
}

function normalizeBarcode(value) {
  return value.trim();
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
