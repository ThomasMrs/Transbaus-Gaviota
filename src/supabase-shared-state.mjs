const SUPABASE_URL = "https://znbcnahjvtdndttqjpec.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_L7q8_g_9SavtGfkNlPjK6Q_imFwj6_N";
const SHARED_STATE_TABLE = "shared_state";
const SHARED_STATE_ROW_ID = "global";
const DEFAULT_PAGE_ID = "global";

export function createSharedStateStore(options = {}) {
  const pageId = normalizePageId(options.pageId || DEFAULT_PAGE_ID);
  const pageTitle = normalizePageTitle(options.pageTitle || "");
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
        pages: listPages(sharedRecord.payload, sharedRecord.updatedAt),
      };
    },

    async saveStateRecord(payload) {
      const sharedRecord = await fetchSharedRecord(client);
      const fallbackUpdatedAt = new Date().toISOString();
      const nextPayload = mergePageState(sharedRecord.payload, pageId, payload, {
        title: pageTitle,
        updatedAt: fallbackUpdatedAt,
      });
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
        pages: listPages(nextPayload, String(data?.updated_at || fallbackUpdatedAt)),
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

function mergePageState(sharedPayload, pageId, nextPagePayload, options = {}) {
  const container = normalizeWorkspaceContainer(sharedPayload);
  const existingMeta = normalizePageMeta(container.pageMeta?.[pageId], pageId, pageId === DEFAULT_PAGE_ID ? "Page principale" : "");
  const title = normalizePageTitle(options.title || existingMeta.title || "");
  const updatedAt = String(options.updatedAt || new Date().toISOString());

  return {
    version: 3,
    pages: {
      ...container.pages,
      [pageId]: nextPagePayload,
    },
    pageMeta: {
      ...container.pageMeta,
      [pageId]: {
        id: pageId,
        title: title || existingMeta.title || defaultPageTitle(pageId),
        createdAt: existingMeta.createdAt || updatedAt,
        updatedAt,
      },
    },
  };
}

function normalizeWorkspaceContainer(sharedPayload) {
  if (isWorkspaceContainer(sharedPayload)) {
    return {
      version: Number(sharedPayload.version) >= 3 ? Number(sharedPayload.version) : 3,
      pages: { ...sharedPayload.pages },
      pageMeta: normalizeWorkspacePageMeta(sharedPayload.pageMeta),
    };
  }

  if (isAppStatePayload(sharedPayload)) {
    return {
      version: 3,
      pages: {
        [DEFAULT_PAGE_ID]: sharedPayload,
      },
      pageMeta: {
        [DEFAULT_PAGE_ID]: {
          id: DEFAULT_PAGE_ID,
          title: "Page principale",
          createdAt: "",
          updatedAt: "",
        },
      },
    };
  }

  return {
    version: 3,
    pages: {},
    pageMeta: {},
  };
}

function listPages(sharedPayload, fallbackUpdatedAt = "") {
  const container = normalizeWorkspaceContainer(sharedPayload);
  const pageIds = new Set([
    ...Object.keys(container.pages),
    ...Object.keys(container.pageMeta),
  ]);

  return [...pageIds]
    .map((pageId) => {
      const state = isAppStatePayload(container.pages[pageId]) ? container.pages[pageId] : null;
      const meta = normalizePageMeta(container.pageMeta[pageId], pageId, pageId === DEFAULT_PAGE_ID ? "Page principale" : "");
      return {
        id: pageId,
        title: meta.title || defaultPageTitle(pageId),
        createdAt: meta.createdAt || "",
        updatedAt: meta.updatedAt || fallbackUpdatedAt || "",
        parcelsCount: Array.isArray(state?.parcels) ? state.parcels.length : 0,
        baquesCount: Array.isArray(state?.baques) ? state.baques.length : 0,
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

function normalizeWorkspacePageMeta(pageMeta) {
  if (!pageMeta || typeof pageMeta !== "object" || Array.isArray(pageMeta)) {
    return {};
  }

  return Object.fromEntries(
    Object.entries(pageMeta)
      .map(([pageId, meta]) => [pageId, normalizePageMeta(meta, pageId)])
      .filter((entry) => entry[1]),
  );
}

function normalizePageMeta(meta, pageId, fallbackTitle = "") {
  return {
    id: normalizePageId(meta?.id || pageId || ""),
    title: normalizePageTitle(meta?.title || fallbackTitle || ""),
    createdAt: String(meta?.createdAt || ""),
    updatedAt: String(meta?.updatedAt || ""),
  };
}

function normalizePageTitle(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function defaultPageTitle(pageId) {
  return pageId === DEFAULT_PAGE_ID ? "Page principale" : pageId.replace(/[-_]+/g, " ");
}
