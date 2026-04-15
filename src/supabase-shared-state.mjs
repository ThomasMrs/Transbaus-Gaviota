const SUPABASE_URL = "https://znbcnahjvtdndttqjpec.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_L7q8_g_9SavtGfkNlPjK6Q_imFwj6_N";
const SUPABASE_ACCESS_EMAIL = "site-access@transbaus.local";
const SHARED_STATE_TABLE = "shared_state";
const DEFAULT_PAGE_ID = "global";
const APP_STATE_COLLECTION_KEYS = [
  "baques",
  "parcels",
  "smallParcelScans",
  "deliveryNotes",
  "destinationRules",
];

export function createSharedStateStore(options = {}) {
  const pageId = normalizePageId(options.pageId || DEFAULT_PAGE_ID);
  const resolvePageTitle = () => normalizePageTitle(
    typeof options.pageTitle === "function" ? options.pageTitle() : options.pageTitle || "",
  );
  const createClient = globalThis.supabase?.createClient;
  if (typeof createClient !== "function") {
    throw new Error("supabase-unavailable");
  }

  const client = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
    auth: {
      autoRefreshToken: true,
      detectSessionInUrl: false,
      persistSession: true,
      storageKey: "transbaus-gaviota-auth-v1",
    },
  });
  let lastFetchedState = buildEmptyAppState();

  return {
    async getAccessSession() {
      const { data, error } = await client.auth.getSession();
      if (error) {
        throw error;
      }

      return data.session || null;
    },

    subscribeAccessState(listener) {
      const callback = typeof listener === "function" ? listener : () => {};
      const {
        data: { subscription },
      } = client.auth.onAuthStateChange((_event, session) => {
        if (!session) {
          lastFetchedState = buildEmptyAppState();
        }
        callback(session || null);
      });

      return () => {
        subscription.unsubscribe();
      };
    },

    async signInWithPassword(password) {
      const { data, error } = await client.auth.signInWithPassword({
        email: SUPABASE_ACCESS_EMAIL,
        password: String(password || ""),
      });

      if (error) {
        throw error;
      }

      return data.session || null;
    },

    async signOut() {
      const { error } = await client.auth.signOut();
      if (error) {
        throw error;
      }

      lastFetchedState = buildEmptyAppState();
    },

    async fetchStateRecord() {
      const [pageRecord, pages] = await Promise.all([
        fetchPageRecord(client, pageId),
        fetchPages(client),
      ]);
      const state = pageRecord ? normalizeAppStatePayload(pageRecord.payload) : null;
      lastFetchedState = state ? normalizeAppStatePayload(state) : buildEmptyAppState();
      return {
        state,
        updatedAt: pageRecord?.updatedAt || "",
        pages,
      };
    },

    async saveStateRecord(payload) {
      const incomingState = normalizeAppStatePayload(payload);
      const currentRecord = await fetchPageRecord(client, pageId);
      const mergedState = mergeAppStates(
        normalizeAppStatePayload(currentRecord?.payload),
        incomingState,
        lastFetchedState,
      );
      const fallbackUpdatedAt = new Date().toISOString();
      const updatedAt = await upsertPageRecord(client, {
        id: pageId,
        title: resolvePageTitle() || defaultPageTitle(pageId),
        createdAt: currentRecord?.createdAt || fallbackUpdatedAt,
        updatedAt: fallbackUpdatedAt,
        payload: mergedState,
      });

      lastFetchedState = mergedState;
      return {
        updatedAt,
        pages: await fetchPages(client),
      };
    },

    async renamePage(targetPageId, nextTitle) {
      const normalizedPageId = normalizePageId(targetPageId);
      if (normalizedPageId === DEFAULT_PAGE_ID) {
        throw createWorkspaceError("workspace-primary-protected", "La page principale ne peut pas etre renommee.");
      }

      const existingRecord = await fetchPageRecord(client, normalizedPageId);
      if (!existingRecord) {
        throw createWorkspaceError("workspace-page-missing", "La page a renommer est introuvable.");
      }

      const fallbackUpdatedAt = new Date().toISOString();
      const { error } = await client
        .from(SHARED_STATE_TABLE)
        .update({
          title: normalizePageTitle(nextTitle) || existingRecord.title || defaultPageTitle(normalizedPageId),
          updated_at: fallbackUpdatedAt,
        })
        .eq("id", normalizedPageId);

      if (error) {
        throw error;
      }

      return {
        updatedAt: fallbackUpdatedAt,
        pages: await fetchPages(client),
      };
    },

    async deletePage(targetPageId) {
      const normalizedPageId = normalizePageId(targetPageId);
      if (normalizedPageId === DEFAULT_PAGE_ID) {
        throw createWorkspaceError("workspace-primary-protected", "La page principale ne peut pas etre supprimee.");
      }

      const existingRecord = await fetchPageRecord(client, normalizedPageId);
      if (!existingRecord) {
        throw createWorkspaceError("workspace-page-missing", "La page a supprimer est introuvable.");
      }

      const { error } = await client
        .from(SHARED_STATE_TABLE)
        .delete()
        .eq("id", normalizedPageId);

      if (error) {
        throw error;
      }

      if (normalizedPageId === pageId) {
        lastFetchedState = buildEmptyAppState();
      }

      return {
        updatedAt: new Date().toISOString(),
        pages: await fetchPages(client),
      };
    },
  };
}

async function fetchPageRecord(client, pageId) {
  const { data, error } = await client
    .from(SHARED_STATE_TABLE)
    .select("id, title, created_at, updated_at, payload")
    .eq("id", pageId)
    .maybeSingle();

  if (error) {
    throw error;
  }

  if (!data) {
    return null;
  }

  return {
    id: normalizePageId(data.id),
    title: normalizePageTitle(data.title || defaultPageTitle(data.id)),
    createdAt: String(data.created_at || ""),
    updatedAt: String(data.updated_at || ""),
    payload: data.payload || null,
  };
}

async function fetchPages(client) {
  const { data, error } = await client
    .from(SHARED_STATE_TABLE)
    .select("id, title, created_at, updated_at, payload");

  if (error) {
    throw error;
  }

  return (Array.isArray(data) ? data : [])
    .map((row) => {
      const state = normalizeAppStatePayload(row.payload);
      return {
        id: normalizePageId(row.id),
        title: normalizePageTitle(row.title || defaultPageTitle(row.id)),
        createdAt: String(row.created_at || ""),
        updatedAt: String(row.updated_at || ""),
        parcelsCount: state.parcels.length + countSmallParcelScans(state.smallParcelScans),
        baquesCount: state.baques.length,
      };
    })
    .sort((left, right) => {
      const rightTime = Date.parse(right.updatedAt || right.createdAt || "") || 0;
      const leftTime = Date.parse(left.updatedAt || left.createdAt || "") || 0;
      if (rightTime !== leftTime) {
        return rightTime - leftTime;
      }

      return left.title.localeCompare(right.title, "fr", { sensitivity: "base" });
    });
}

async function upsertPageRecord(client, record) {
  const { data, error } = await client
    .from(SHARED_STATE_TABLE)
    .upsert({
      id: normalizePageId(record.id),
      title: normalizePageTitle(record.title || defaultPageTitle(record.id)),
      created_at: record.createdAt || record.updatedAt,
      updated_at: record.updatedAt,
      payload: normalizeAppStatePayload(record.payload),
    }, {
      onConflict: "id",
    })
    .select("updated_at")
    .single();

  if (error) {
    throw error;
  }

  return String(data?.updated_at || record.updatedAt);
}

function normalizePageId(value) {
  const normalizedValue = String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-_]/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^-+|-+$/g, "");

  return normalizedValue || DEFAULT_PAGE_ID;
}

function normalizePageTitle(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function defaultPageTitle(pageId) {
  return normalizePageId(pageId) === DEFAULT_PAGE_ID ? "Page principale" : String(pageId || "").replace(/[-_]+/g, " ");
}

function isAppStatePayload(payload) {
  return Boolean(
    payload
    && typeof payload === "object"
    && APP_STATE_COLLECTION_KEYS.every((key) => Array.isArray(payload[key] || []))
    && Array.isArray(payload.baques)
    && Array.isArray(payload.parcels)
    && Array.isArray(payload.deliveryNotes)
    && Array.isArray(payload.destinationRules),
  );
}

function buildEmptyAppState() {
  return {
    baques: [],
    parcels: [],
    smallParcelScans: [],
    deliveryNotes: [],
    destinationRules: [],
  };
}

function normalizeAppStatePayload(payload) {
  if (!isAppStatePayload(payload)) {
    return buildEmptyAppState();
  }

  return {
    baques: normalizeEntityCollection(payload.baques),
    parcels: normalizeEntityCollection(payload.parcels),
    smallParcelScans: normalizeEntityCollection(payload.smallParcelScans),
    deliveryNotes: normalizeEntityCollection(payload.deliveryNotes),
    destinationRules: normalizeEntityCollection(payload.destinationRules),
  };
}

function normalizeEntityCollection(items) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items
    .filter((item) => item && typeof item === "object" && item.id)
    .map((item) => JSON.parse(JSON.stringify(item)));
}

function mergeAppStates(remoteState, localState, baseState) {
  const normalizedRemote = normalizeAppStatePayload(remoteState);
  const normalizedLocal = normalizeAppStatePayload(localState);
  const normalizedBase = normalizeAppStatePayload(baseState);

  return Object.fromEntries(
    APP_STATE_COLLECTION_KEYS.map((key) => [
      key,
      mergeEntityCollection(normalizedRemote[key], normalizedLocal[key], normalizedBase[key]),
    ]),
  );
}

function mergeEntityCollection(remoteItems, localItems, baseItems) {
  const remoteMap = toEntityMap(remoteItems);
  const localMap = toEntityMap(localItems);
  const baseMap = toEntityMap(baseItems);
  const mergedMap = new Map();
  const allIds = new Set([
    ...remoteMap.keys(),
    ...localMap.keys(),
    ...baseMap.keys(),
  ]);

  allIds.forEach((id) => {
    const mergedEntity = resolveMergedEntity(remoteMap.get(id), localMap.get(id), baseMap.get(id));
    if (mergedEntity) {
      mergedMap.set(id, mergedEntity);
    }
  });

  const localOrder = Array.isArray(localItems) ? localItems.map((item) => String(item.id)) : [];
  const remoteOrder = Array.isArray(remoteItems) ? remoteItems.map((item) => String(item.id)) : [];
  const mergedOrder = [
    ...localOrder,
    ...remoteOrder.filter((id) => !localOrder.includes(id)),
  ];

  return [
    ...mergedOrder
      .map((id) => mergedMap.get(id))
      .filter(Boolean),
    ...[...mergedMap.entries()]
      .filter(([id]) => !mergedOrder.includes(id))
      .map((entry) => entry[1]),
  ];
}

function resolveMergedEntity(remoteEntity, localEntity, baseEntity) {
  if (!remoteEntity && !localEntity) {
    return null;
  }

  if (baseEntity) {
    if (!localEntity) {
      return null;
    }

    if (!remoteEntity) {
      return hasEntityChangedSinceBase(localEntity, baseEntity) ? localEntity : null;
    }
  }

  if (!remoteEntity) {
    return localEntity || null;
  }

  if (!localEntity) {
    return remoteEntity || null;
  }

  const remoteTimestamp = getEntityTimestamp(remoteEntity);
  const localTimestamp = getEntityTimestamp(localEntity);
  if (remoteTimestamp && localTimestamp) {
    return Date.parse(localTimestamp) >= Date.parse(remoteTimestamp) ? localEntity : remoteEntity;
  }

  return localEntity;
}

function hasEntityChangedSinceBase(entity, baseEntity) {
  const currentTimestamp = getEntityTimestamp(entity);
  const baseTimestamp = getEntityTimestamp(baseEntity);
  if (currentTimestamp && baseTimestamp && currentTimestamp !== baseTimestamp) {
    return true;
  }

  return JSON.stringify(entity) !== JSON.stringify(baseEntity);
}

function getEntityTimestamp(entity) {
  if (!entity || typeof entity !== "object") {
    return "";
  }

  return String(entity.updatedAt || entity.analyzedAt || entity.importedAt || entity.createdAt || "");
}

function toEntityMap(items) {
  return new Map(
    (Array.isArray(items) ? items : [])
      .filter((item) => item && typeof item === "object" && item.id)
      .map((item) => [String(item.id), item]),
  );
}

function createWorkspaceError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function countSmallParcelScans(scans) {
  if (!Array.isArray(scans)) {
    return 0;
  }

  return scans.reduce((total, scan) => total + Math.max(1, Number(scan?.quantity || 1)), 0);
}

export const __testables = {
  buildEmptyAppState,
  mergeAppStates,
  normalizeAppStatePayload,
};
