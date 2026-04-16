import {
  normalizeFreeText,
  normalizeStoredDate,
} from "./shared.mjs";
import {
  getParcelCityLabel,
  getParcelCommandNumber,
  normalizeBarcode,
  normalizeCommandNumber,
  sanitizeDestination,
} from "./parcel-utils.mjs";

export function normalizeDeliveryNote(note) {
  if (!note || !note.id || !note.name || !note.importedAt) {
    return null;
  }

  const importedAt = normalizeStoredDate(note.importedAt, new Date().toISOString());

  return {
    id: String(note.id),
    name: normalizeFreeText(String(note.name)),
    size: Number(note.size || 0),
    importedAt,
    updatedAt: normalizeStoredDate(note.updatedAt || note.analysis?.analyzedAt || importedAt, importedAt),
    importedByEmail: normalizeFreeText(note.importedByEmail || "").toLowerCase(),
    importedByLabel: normalizeFreeText(note.importedByLabel || ""),
    updatedByEmail: normalizeFreeText(note.updatedByEmail || note.importedByEmail || "").toLowerCase(),
    updatedByLabel: normalizeFreeText(note.updatedByLabel || note.importedByLabel || ""),
    analysis: normalizeDeliveryNoteAnalysis(note.analysis),
  };
}

export function looksLikePdf(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name || "");
}

export function getDeliveryNoteErrorMessage(error, fallbackMessage = "Impossible d'analyser ce PDF.") {
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

export function formatFileSize(size) {
  const bytes = Number(size || 0);
  if (bytes >= 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(1)} Mo`;
  }

  if (bytes >= 1024) {
    return `${Math.round(bytes / 1024)} Ko`;
  }

  return `${bytes} o`;
}

export function parseDeliveryNoteText(text) {
  const lines = String(text ?? "")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => normalizePdfOcrLine(line))
    .filter(Boolean);
  const structuredEntries = dedupeDeliveryEntries(parseStructuredDeliveryNoteLines(lines));
  const legacyEntries = dedupeDeliveryEntries(parseLegacyDeliveryNoteLines(lines));

  return structuredEntries.length >= legacyEntries.length ? structuredEntries : legacyEntries;
}

export function compareDeliveryNoteEntries(entries, parcels) {
  const incomparableParcelsCount = countIncomparableParcels(parcels);
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

  const registeredCommandCounts = buildRegisteredCommandCounts(parcels);
  const registeredCommandInfo = buildRegisteredCommandInfo(parcels);
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

export function countIncomparableParcels(parcels) {
  return parcels.filter((parcel) => !getParcelCommandNumber(parcel)).length;
}

export function buildDeliveryEntryDestinationLabel(entry) {
  const city = sanitizeDestination(entry?.city || "");
  const postalCode = extractDeliveryEntryPostalCode(entry);
  if (postalCode && city) {
    return `${postalCode} ${city}`;
  }

  return city || postalCode || "";
}

function normalizeDeliveryNoteAnalysis(analysis) {
  if (!analysis || !Array.isArray(analysis.missingEntries)) {
    return null;
  }

  const entries = Array.isArray(analysis.entries)
    ? analysis.entries
      .map((entry) => normalizeDeliveryNoteEntry(entry))
      .filter((entry) => entry.commandNumber)
    : [];

  return {
    totalEntries: Number(analysis.totalEntries || 0),
    totalExpectedCount: Number(analysis.totalExpectedCount || 0),
    totalRegisteredCount: Number(analysis.totalRegisteredCount || 0),
    totalMissingCount: Number(analysis.totalMissingCount || 0),
    incomparableParcelsCount: Number(analysis.incomparableParcelsCount || 0),
    parseError: normalizeFreeText(analysis.parseError || ""),
    entries,
    missingEntries: analysis.missingEntries
      .map((entry) => ({
        ...normalizeDeliveryNoteEntry(entry),
        registeredCount: Number(entry.registeredCount || 0),
        missingCount: Number(entry.missingCount || 0),
      }))
      .filter((entry) => entry.commandNumber),
    analyzedAt: normalizeStoredDate(analysis.analyzedAt || "", ""),
  };
}

function normalizeDeliveryNoteEntry(entry) {
  return {
    commandNumber: normalizeCommandNumber(entry?.commandNumber || ""),
    expectedCount: Number(entry?.expectedCount || 1),
    client: normalizeFreeText(entry?.client || ""),
    city: normalizeFreeText(entry?.city || ""),
    rawContext: normalizeFreeText(entry?.rawContext || ""),
  };
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

function normalizePdfOcrLine(value) {
  return normalizeFreeText(
    String(value ?? "")
      .replaceAll("\u0000", " ")
      .replace(/[|]/g, "I")
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'")
      .replace(/[^\S\r\n]+/g, " "),
  );
}

export function normalizeDeliveryTextLine(value) {
  return normalizeFreeText(
    String(value ?? "")
      .replace(/\bO(?=\d)/g, "0")
      .replace(/(?<=\d)O\b/g, "0")
      .replace(/[|]/g, "I"),
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

function buildRegisteredCommandCounts(parcels) {
  const groupedCounts = parcels.reduce((map, parcel) => {
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

function buildRegisteredCommandInfo(parcels) {
  const groupedInfo = parcels.reduce((map, parcel) => {
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

function getMostCommonDeliveryValue(counts) {
  return [...counts.entries()]
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0], "fr", { sensitivity: "base" }))[0]?.[0] || "";
}

function extractDeliveryEntryPostalCode(entry) {
  const rawContext = normalizeDeliveryTextLine(entry?.rawContext || "");
  const match = rawContext.match(/\b\d{5}\b/);
  return match ? match[0] : "";
}

function looksLikeDeliveryDate(value) {
  return /^(?:19|20)\d{6}$/.test(value) || /^(?:0[1-9]|[12]\d|3[01])(?:0[1-9]|1[0-2])\d{2}$/.test(value);
}
