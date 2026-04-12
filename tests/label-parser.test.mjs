import test from "node:test";
import assert from "node:assert/strict";

import { parseLabelText } from "../src/label-parser.mjs";

test("parseLabelText extracts parcel data from OCR text", () => {
  const parsed = parseLabelText(`
CLIENT MENUISERIE VIDAL
ADRESSE 47500 SAINT-VITE
DESCRIPTION Systeme coffre
ROUTE R4 TRANSBAUS 4 - Zone 47 Lot-et-Garonne
REF BERNERD
DATE 25/03/2026
COMMANDE 063619
25/03/2026
49,41 Kg
1/2
R447500
  `);

  assert.equal(parsed.commandNumber, "063619");
  assert.equal(parsed.routeCode, "R447500");
  assert.equal(parsed.destination, "47500 SAINT-VITE");
  assert.equal(parsed.client, "MENUISERIE VIDAL");
  assert.match(parsed.reference, /^BERNERD\b/);
  assert.equal(parsed.weight, "49,41 Kg");
  assert.equal(parsed.packageIndex, "1/2");
});
