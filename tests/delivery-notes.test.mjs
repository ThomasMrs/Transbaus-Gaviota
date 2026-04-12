import test from "node:test";
import assert from "node:assert/strict";

import {
  compareDeliveryNoteEntries,
  parseDeliveryNoteText,
} from "../src/delivery-notes.mjs";

test("parseDeliveryNoteText extracts structured delivery entries", () => {
  const entries = parseDeliveryNoteText(`
CODE POSTAL VILLE CLIENT
MENUISERIE VIDAL
47500 SAINT-VITE
063619 25/03/2026 2,00
COMMANDE 063619
*063619002*
  `);

  assert.equal(entries.length, 1);
  assert.equal(entries[0].commandNumber, "063619");
  assert.equal(entries[0].expectedCount, 2);
  assert.equal(entries[0].client, "MENUISERIE VIDAL");
  assert.match(entries[0].city, /SAINT-VITE/);
});

test("compareDeliveryNoteEntries reports missing packages and incomparable parcels", () => {
  const entries = [
    {
      commandNumber: "063619",
      expectedCount: 2,
      client: "MENUISERIE VIDAL",
      city: "SAINT-VITE",
      rawContext: "47500 SAINT-VITE | COMMANDE 063619",
    },
  ];
  const parcels = [
    { commandNumber: "063619", barcode: "063619001", packageIndex: "1/2", destination: "47500 SAINT-VITE", client: "MENUISERIE VIDAL" },
    { commandNumber: "", barcode: "", packageIndex: "", destination: "47000 AGEN", client: "Client sans commande" },
  ];

  const analysis = compareDeliveryNoteEntries(entries, parcels);

  assert.equal(analysis.totalEntries, 1);
  assert.equal(analysis.totalExpectedCount, 2);
  assert.equal(analysis.totalRegisteredCount, 1);
  assert.equal(analysis.totalMissingCount, 1);
  assert.equal(analysis.incomparableParcelsCount, 1);
  assert.equal(analysis.missingEntries[0].missingCount, 1);
});
