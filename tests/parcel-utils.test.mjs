import test from "node:test";
import assert from "node:assert/strict";

import {
  findExistingParcel,
  formatRouteCodeForDisplay,
  normalizeParcelData,
  parseParcelWeightKg,
} from "../src/parcel-utils.mjs";

test("normalizeParcelData derives command number and route code", () => {
  const parcel = normalizeParcelData({
    barcode: "063619001",
    routeCode: "",
    destination: "47500 Saint-Vite",
    client: "Menuiserie Vidal",
    routeLabel: "R4 TRANSBAUS 4 - Zone 47 Lot-et-Garonne",
    shippingDate: "25/03/2026",
    weight: "49,41 Kg",
    packageIndex: "1/2",
  });

  assert.equal(parcel.commandNumber, "063619");
  assert.equal(parcel.routeCode, "R447500");
  assert.equal(parcel.destination, "47500 Saint-Vite");
  assert.equal(parcel.packageIndex, "1/2");
});

test("findExistingParcel matches the exact package within a command", () => {
  const parcels = [
    { commandNumber: "063619", barcode: "063619001", packageIndex: "1/2" },
    { commandNumber: "063619", barcode: "063619002", packageIndex: "2/2" },
  ];

  const match = findExistingParcel(parcels, {
    commandNumber: "063619",
    barcode: "",
    packageIndex: "2/2",
  });

  assert.ok(match);
  assert.equal(match.packageIndex, "2/2");
});

test("formatRouteCodeForDisplay and parseParcelWeightKg keep operator-facing formatting", () => {
  assert.equal(formatRouteCodeForDisplay("R447500"), "R4 47 500");
  assert.equal(parseParcelWeightKg("49,41 Kg"), 49.41);
});
