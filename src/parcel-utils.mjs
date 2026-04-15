import {
  normalizeFreeText,
  stripDiacritics,
} from "./shared.mjs";

export function normalizeDestination(value) {
  return sanitizeDestination(value);
}

export function normalizeRouteCode(value) {
  return normalizeFreeText(value).replace(/\s+/g, "").toUpperCase();
}

export function parseDestinationRulePatterns(value) {
  const rawValues = Array.isArray(value)
    ? value
    : String(value || "")
      .replace(/,/g, "\n")
      .split("\n");

  return [...new Set(rawValues.map((entry) => normalizeDestinationRulePattern(entry)).filter(Boolean))];
}

export function normalizeDestinationRulePattern(value) {
  const rawValue = normalizeFreeText(String(value || ""));
  const directPostalCode = rawValue.match(/^\d{5}$/)?.[0];
  if (directPostalCode) {
    return directPostalCode;
  }

  return normalizeDestinationRuleText(rawValue);
}

export function normalizeDestinationRuleText(value) {
  return stripDiacritics(sanitizeDestination(String(value || "")).toUpperCase())
    .replace(/[^0-9A-Z\s-]/g, " ")
    .replace(/[-/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeParcelData(parcel) {
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
    measuredDimensions: normalizeFreeText(parcel.measuredDimensions || ""),
    packageIndex,
  };
}

export function getParcelIdentifier(parcel) {
  return getParcelCommandNumber(parcel) || parcel.barcode || formatRouteCodeForDisplay(parcel.routeCode) || "Sans code";
}

export function findExistingParcel(parcels, parcelData) {
  const barcode = normalizeBarcode(parcelData.barcode || "");
  const packageIndex = normalizeFreeText(parcelData.packageIndex || "");
  const commandNumber = getParcelCommandNumber(parcelData);

  if (commandNumber && packageIndex) {
    const exactCommandMatch = parcels.find(
      (parcel) => getParcelCommandNumber(parcel) === commandNumber && normalizeFreeText(parcel.packageIndex || "") === packageIndex,
    );
    if (exactCommandMatch) {
      return exactCommandMatch;
    }
  }

  if (barcode && packageIndex) {
    const exactBarcodeMatch = parcels.find(
      (parcel) => normalizeBarcode(parcel.barcode || "") === barcode && normalizeFreeText(parcel.packageIndex || "") === packageIndex,
    );
    if (exactBarcodeMatch) {
      return exactBarcodeMatch;
    }
  }

  if (commandNumber && !packageIndex) {
    const commandMatches = parcels.filter(
      (parcel) => getParcelCommandNumber(parcel) === commandNumber && !normalizeFreeText(parcel.packageIndex || ""),
    );
    if (commandMatches.length === 1) {
      return commandMatches[0];
    }
  }

  if (barcode && !packageIndex) {
    const barcodeMatches = parcels.filter(
      (parcel) => normalizeBarcode(parcel.barcode || "") === barcode && !normalizeFreeText(parcel.packageIndex || ""),
    );
    if (barcodeMatches.length === 1) {
      return barcodeMatches[0];
    }
  }

  return null;
}

export function sanitizeDestination(value) {
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

export function cleanDestinationSegment(segment) {
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

export function sanitizePackageIndex(packageIndex, shippingDate) {
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

export function reconcileDestination(destination, routeCode) {
  const routePostalCode = routeCode.match(/(\d{5})$/)?.[1] || "";
  const destinationPostalCode = extractPostalCode(destination);

  if (!routePostalCode || !destinationPostalCode || routePostalCode === destinationPostalCode) {
    return destination;
  }

  return normalizeFreeText(destination.replace(destinationPostalCode, routePostalCode));
}

export function reconcileRouteCode(routeCode, routeLabel, destination) {
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

export function reconcileBarcode(barcode, destination, routeCode) {
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

export function reconcileCommandNumber(commandNumber, barcode) {
  return normalizeCommandNumber(commandNumber) || deriveCommandNumberFromBarcode(barcode);
}

export function extractPostalCode(destination) {
  return String(destination).match(/\b\d{5}\b/)?.[0] || "";
}

export function getDestinationShortLabel(destination) {
  const cleaned = sanitizeDestination(destination);
  const match = cleaned.match(/\b\d{5}\s+[A-ZÀ-ÿ-]+(?:\s+[A-ZÀ-ÿ-]+){0,2}/u);
  return match ? normalizeFreeText(match[0]) : cleaned;
}

export function normalizeBarcode(value) {
  return String(value ?? "").trim();
}

export function parseParcelWeightKg(parcel) {
  const rawWeight = typeof parcel === "string" ? parcel : parcel?.weight || "";
  const match = String(rawWeight).match(/(\d+(?:[.,]\d+)?)/);
  if (!match) {
    return null;
  }

  const parsedWeight = Number.parseFloat(match[1].replace(",", "."));
  return Number.isFinite(parsedWeight) ? parsedWeight : null;
}

export function getParcelHandlingFactor(weightKg) {
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

export function getSortingEffortLabel(totalEffort) {
  if (totalEffort >= 18) {
    return "Effort soutenu";
  }

  if (totalEffort >= 8) {
    return "Effort moyen";
  }

  return "Effort leger";
}

export function getSortingHandlingAdvice(summary) {
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

export function formatKnownWeightSummary(totalWeightKg, unknownWeightCount = 0) {
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

export function normalizeCommandNumber(value) {
  const normalizedValue = normalizeBarcode(String(value || ""));
  return /^\d{5,8}$/.test(normalizedValue) ? normalizedValue : "";
}

export function deriveCommandNumberFromBarcode(barcode) {
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

export function getParcelCommandNumber(parcel) {
  return reconcileCommandNumber(parcel?.commandNumber || "", parcel?.barcode || "");
}

export function getParcelCityLabel(parcel) {
  const destination = sanitizeDestination(parcel.destination || "");
  const match = destination.match(/\b\d{5}\s+(.+)$/);
  return normalizeFreeText(match ? match[1] : destination);
}

export function formatRouteCodeForDisplay(routeCode) {
  const normalized = normalizeRouteCode(routeCode || "");
  const match = normalized.match(/^R(\d+)(\d{5})$/);
  if (!match) {
    return normalized;
  }

  const routePrefix = `R${match[1]}`;
  const postalCode = match[2];
  return `${routePrefix} ${postalCode.slice(0, 2)} ${postalCode.slice(2)}`;
}
