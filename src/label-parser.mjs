import { normalizeFreeText } from "./shared.mjs";
import {
  normalizeParcelData,
  normalizeRouteCode,
  sanitizeDestination,
} from "./parcel-utils.mjs";

export function parseLabelText(text) {
  const lines = String(text ?? "")
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

  const directMatch = String(text ?? "").match(/(?:N\W*COMMANDE|COMMANDE)[^\dA-Z]*([0-9]{5,10})/i);
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

  const fromText = String(text ?? "").match(routeCodeRegex);
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
  const match = String(text ?? "").match(/\b\d{2}\/\d{2}\/\d{4}\b/);
  return match ? match[0] : "";
}

function extractWeight(text) {
  const match = String(text ?? "").match(/\b\d{1,3}[.,]\d{1,2}\s*Kg\b/i);
  return match ? normalizeFreeText(match[0].replace(/\s+/g, " ")) : "";
}

function extractPackageIndex(lines, text) {
  const lineMatch = lines.find((line) => /^\d{1,2}\s*\/\s*\d{1,2}$/.test(line));
  if (lineMatch) {
    return lineMatch.replace(/\s+/g, "");
  }

  const match = String(text ?? "").match(/\b\d{1,2}\s*\/\s*\d{1,2}\b(?!\s*\/)/);
  return match ? match[0].replace(/\s+/g, "") : "";
}

function joinSectionLines(lines, separator = " ") {
  return normalizeFreeText(lines.filter(Boolean).join(separator));
}

function normalizeOcrLine(value) {
  return normalizeFreeText(
    String(value ?? "")
      .replace(/[|]/g, "I")
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'"),
  );
}
