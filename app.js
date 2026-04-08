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
  importingBarcode: false,
};
const ocr = {
  worker: null,
  busy: false,
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
  ui.scanLabelBtn = document.querySelector("#scanLabelBtn");
  ui.labelImageInput = document.querySelector("#labelImageInput");
  ui.barcodeImageInput = document.querySelector("#barcodeImageInput");
  ui.ocrStatus = document.querySelector("#ocrStatus");
  ui.barcodeStatus = document.querySelector("#barcodeStatus");
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
  ui.barcodeFileReader = document.querySelector("#barcodeFileReader");
}

function bindEvents() {
  ui.parcelForm.addEventListener("submit", handleParcelSubmit);
  ui.baqueForm.addEventListener("submit", handleBaqueSubmit);
  ui.searchInput.addEventListener("input", renderSearchResults);
  ui.openScannerBtn.addEventListener("click", openScanner);
  ui.importBarcodeBtn.addEventListener("click", openBarcodeImagePicker);
  ui.scanLabelBtn.addEventListener("click", openLabelScanner);
  ui.barcodeImageInput.addEventListener("change", handleBarcodeImageChange);
  ui.labelImageInput.addEventListener("change", handleLabelImageChange);
  ui.closeScannerBtn.addEventListener("click", closeScanner);
  ui.scannerModal.addEventListener("click", handleModalClick);
  ui.baquesGrid.addEventListener("click", handleBaqueGridClick);
  ui.baquesGrid.addEventListener("change", handleBaqueGridChange);
  window.addEventListener("beforeunload", () => {
    void stopScanner();
    void stopOcrWorker();
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
      .map((parcel) => normalizeParcelData({
        id: parcel.id || createId(),
        barcode: String(parcel.barcode || "").trim(),
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
      .filter((parcel) => parcel.routeCode || parcel.barcode || parcel.destination);

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
  const detailLines = [
    `Destination <strong>${escapeHtml(displayDestination)}</strong>`,
    parcel.client ? `Client : ${escapeHtml(parcel.client)}` : "",
    parcel.routeCode ? `Numero destination : ${escapeHtml(parcel.routeCode)}` : "",
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
  const tagLabel = parcel.routeCode || getDestinationShortLabel(displayDestination) || "Colis";
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
            ${parcel.routeCode ? `<span><strong>Numero destination :</strong> ${escapeHtml(parcel.routeCode)}</span>` : ""}
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

function openLabelScanner() {
  if (ocr.busy) {
    return;
  }

  ui.labelImageInput.click();
}

function openBarcodeImagePicker() {
  if (scanner.importingBarcode) {
    return;
  }

  ui.barcodeImageInput.click();
}

async function handleBarcodeImageChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  if (typeof window.Html5Qrcode === "undefined") {
    showToast("La librairie de scan n'a pas pu etre chargee.", "danger");
    ui.barcodeImageInput.value = "";
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
    ui.barcodeImageInput.value = "";
    setBarcodeImportBusy(false);
  }
}

async function handleLabelImageChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  if (typeof window.Tesseract?.createWorker !== "function") {
    showToast("Le module OCR n'est pas disponible.", "danger");
    ui.labelImageInput.value = "";
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
    ui.labelImageInput.value = "";
  }
}

function setOcrBusy(isBusy) {
  ocr.busy = isBusy;
  ui.scanLabelBtn.disabled = isBusy;
  ui.scanLabelBtn.textContent = isBusy ? "Analyse etiquette..." : "Importer photo etiquette";
}

function setBarcodeImportBusy(isBusy) {
  scanner.importingBarcode = isBusy;
  ui.importBarcodeBtn.disabled = isBusy;
  ui.importBarcodeBtn.textContent = isBusy ? "Analyse code-barres..." : "Importer photo code-barres";
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
    const candidates = lines
      .slice(commandIndex, commandIndex + 4)
      .flatMap((line) => [...line.matchAll(/\b\d{5,10}\b/g)].map((match) => match[0]))
      .filter((value, index, array) => array.indexOf(value) === index)
      .sort((left, right) => left.length - right.length);

    const preferred = candidates.find((value) => value.length >= 5 && value.length <= 8);
    if (preferred) {
      return preferred;
    }
  }

  const directMatch = text.match(/(?:N\W*COMMANDE|COMMANDE)[^\dA-Z]*([0-9]{5,10})/i);
  if (directMatch) {
    return directMatch[1].replace(/\s+/g, "");
  }

  return "";
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
  const existing = normalizedParcelData.barcode
    ? state.parcels.find((parcel) => parcel.barcode === normalizedParcelData.barcode)
    : null;

  if (existing) {
    const moved = existing.currentBaqueId !== baqueId;
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
  showToast(`Colis ${getParcelIdentifier(normalizedParcelData)} ajoute dans ${baque.name}.`);
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
  const routeLabel = normalizeFreeText(parcel.routeLabel || "");
  const destination = sanitizeDestination(parcel.destination || "");
  const shippingDate = normalizeFreeText(parcel.shippingDate || "");
  const routeCode = reconcileRouteCode(parcel.routeCode || "", routeLabel, destination);
  const packageIndex = sanitizePackageIndex(parcel.packageIndex || "", shippingDate);

  return {
    ...parcel,
    barcode,
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
  return parcel.destination || parcel.routeCode || "Sans destination";
}

function getParcelDestinationKey(parcel) {
  return getParcelDestinationDisplay(parcel).toUpperCase();
}

function getParcelIdentifier(parcel) {
  return parcel.barcode || parcel.routeCode || "Sans code-barres";
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

  return `<span><strong>Route :</strong> ${escapeHtml(routeCodes[0])}</span>`;
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
