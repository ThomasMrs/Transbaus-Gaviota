import test from "node:test";
import assert from "node:assert/strict";

import { __testables } from "../src/supabase-shared-state.mjs";

const { buildEmptyAppState, mergeAppStates, normalizeAppStatePayload } = __testables;

function createState(overrides = {}) {
  return normalizeAppStatePayload({
    ...buildEmptyAppState(),
    ...overrides,
  });
}

test("mergeAppStates keeps non-conflicting remote and local changes", () => {
  const baseState = createState({
    baques: [
      { id: "b1", name: "Baque 1", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
    parcels: [
      { id: "p1", commandNumber: "063619", currentBaqueId: "b1", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
  });

  const remoteState = createState({
    baques: [
      { id: "b1", name: "Baque 1 Renommee", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:03:00.000Z" },
    ],
    parcels: [
      { id: "p1", commandNumber: "063619", currentBaqueId: "b1", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
      { id: "p2", commandNumber: "063620", currentBaqueId: "b1", createdAt: "2026-04-16T08:04:00.000Z", updatedAt: "2026-04-16T08:04:00.000Z" },
    ],
  });

  const localState = createState({
    baques: [
      { id: "b1", name: "Baque 1", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
    parcels: [
      { id: "p1", commandNumber: "063619", currentBaqueId: "b2", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:05:00.000Z" },
    ],
    destinationRules: [
      { id: "r1", label: "Saint-Vite", patterns: ["saint-vite"], matchMode: "any", preferredBaqueId: "b1", createdAt: "2026-04-16T08:05:00.000Z", updatedAt: "2026-04-16T08:05:00.000Z" },
    ],
  });

  const mergedState = mergeAppStates(remoteState, localState, baseState);

  assert.equal(mergedState.baques[0].name, "Baque 1 Renommee");
  assert.equal(mergedState.parcels.length, 2);
  assert.equal(mergedState.parcels.find((parcel) => parcel.id === "p1")?.currentBaqueId, "b2");
  assert.equal(mergedState.parcels.find((parcel) => parcel.id === "p2")?.commandNumber, "063620");
  assert.equal(mergedState.destinationRules.length, 1);
});

test("mergeAppStates preserves a local deletion relative to the base state", () => {
  const baseState = createState({
    parcels: [
      { id: "p1", commandNumber: "063619", currentBaqueId: "b1", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
  });

  const remoteState = createState({
    parcels: [
      { id: "p1", commandNumber: "063619", currentBaqueId: "b1", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
  });

  const localState = createState({
    parcels: [],
  });

  const mergedState = mergeAppStates(remoteState, localState, baseState);

  assert.equal(mergedState.parcels.length, 0);
});

test("mergeAppStates keeps the newest version when the same entity changed remotely and locally", () => {
  const baseState = createState({
    baques: [
      { id: "b1", name: "Baque 1", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:00:00.000Z" },
    ],
  });

  const remoteState = createState({
    baques: [
      { id: "b1", name: "Baque distante", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:06:00.000Z" },
    ],
  });

  const localState = createState({
    baques: [
      { id: "b1", name: "Baque locale", location: "Zone A", createdAt: "2026-04-16T08:00:00.000Z", updatedAt: "2026-04-16T08:05:00.000Z" },
    ],
  });

  const mergedState = mergeAppStates(remoteState, localState, baseState);

  assert.equal(mergedState.baques[0].name, "Baque distante");
});
