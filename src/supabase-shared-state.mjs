const SUPABASE_URL = "https://znbcnahjvtdndttqjpec.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_L7q8_g_9SavtGfkNlPjK6Q_imFwj6_N";
const SHARED_STATE_TABLE = "shared_state";
const SHARED_STATE_ROW_ID = "global";
const DEFAULT_PAGE_ID = "global";

export function createSharedStateStore(options = {}) {
  const pageId = normalizePageId(options.pageId || DEFAULT_PAGE_ID);
  const createClient = globalThis.supabase?.createClient;
  if (typeof createClient !== "function") {
    throw new Error("supabase-unavailable");
  }

  const client = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
    auth: {
      autoRefreshToken: false,
      detectSessionInUrl: false,
      persistSession: false,
    },
  });

  return {
    async fetchStateRecord() {
      const sharedRecord = await fetchSharedRecord(client);
      return {
        state: extractPageState(sharedRecord.payload, pageId),
        updatedAt: sharedRecord.updatedAt,
      };
    },

    async saveStateRecord(payload) {
      const sharedRecord = await fetchSharedRecord(client);
      const fallbackUpdatedAt = new Date().toISOString();
      const nextPayload = mergePageState(sharedRecord.payload, pageId, payload);
      const { data, error } = await client
        .from(SHARED_STATE_TABLE)
        .upsert({
          id: SHARED_STATE_ROW_ID,
          payload: nextPayload,
          updated_at: fallbackUpdatedAt,
        }, {
          onConflict: "id",
        })
        .select("updated_at")
        .single();

      if (error) {
        throw error;
      }

      return {
        updatedAt: String(data?.updated_at || fallbackUpdatedAt),
      };
    },
  };
}

async function fetchSharedRecord(client) {
  const { data, error } = await client
    .from(SHARED_STATE_TABLE)
    .select("payload, updated_at")
    .eq("id", SHARED_STATE_ROW_ID)
    .maybeSingle();

  if (error) {
    throw error;
  }

  return {
    payload: data?.payload || null,
    updatedAt: String(data?.updated_at || ""),
  };
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

function isAppStatePayload(payload) {
  return Boolean(
    payload
    && typeof payload === "object"
    && Array.isArray(payload.baques)
    && Array.isArray(payload.parcels)
    && Array.isArray(payload.deliveryNotes)
    && Array.isArray(payload.destinationRules),
  );
}

function isWorkspaceContainer(payload) {
  return Boolean(
    payload
    && typeof payload === "object"
    && Number(payload.version) >= 2
    && payload.pages
    && typeof payload.pages === "object"
    && !Array.isArray(payload.pages),
  );
}

function extractPageState(sharedPayload, pageId) {
  if (isAppStatePayload(sharedPayload)) {
    return pageId === DEFAULT_PAGE_ID ? sharedPayload : null;
  }

  if (!isWorkspaceContainer(sharedPayload)) {
    return null;
  }

  const pagePayload = sharedPayload.pages[pageId];
  return isAppStatePayload(pagePayload) ? pagePayload : null;
}

function mergePageState(sharedPayload, pageId, nextPagePayload) {
  const container = normalizeWorkspaceContainer(sharedPayload);
  return {
    version: 2,
    pages: {
      ...container.pages,
      [pageId]: nextPagePayload,
    },
  };
}

function normalizeWorkspaceContainer(sharedPayload) {
  if (isWorkspaceContainer(sharedPayload)) {
    return {
      version: 2,
      pages: { ...sharedPayload.pages },
    };
  }

  if (isAppStatePayload(sharedPayload)) {
    return {
      version: 2,
      pages: {
        [DEFAULT_PAGE_ID]: sharedPayload,
      },
    };
  }

  return {
    version: 2,
    pages: {},
  };
}
