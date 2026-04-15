const SUPABASE_URL = "https://znbcnahjvtdndttqjpec.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_L7q8_g_9SavtGfkNlPjK6Q_imFwj6_N";
const SHARED_STATE_TABLE = "shared_state";
const SHARED_STATE_ROW_ID = "global";
const DEFAULT_PAGE_ID = "global";

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
      autoRefreshToken: false,
      detectSessionInUrl: false,
      persistSession: false,
    },
  });

  return {
    async verifyAccessPassword(password) {
      const { data, error } = await client.rpc("verify_site_access", {
        input_password: String(password || ""),
      });

      if (error) {
        throw error;
      }

      return Boolean(data);
    },

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
        title: resolvePageTitle(),
        updatedAt: fallbackUpdatedAt,
      });
      const updatedAt = await upsertSharedPayload(client, nextPayload, fallbackUpdatedAt);

      return {
        updatedAt,
        pages: listPages(nextPayload, updatedAt),
      };
    },

    async renamePage(targetPageId, nextTitle) {
      const sharedRecord = await fetchSharedRecord(client);
      const fallbackUpdatedAt = new Date().toISOString();
      const nextPayload = renameWorkspacePage(sharedRecord.payload, targetPageId, nextTitle, fallbackUpdatedAt);
      const updatedAt = await upsertSharedPayload(client, nextPayload, fallbackUpdatedAt);
      return {
        updatedAt,
        pages: listPages(nextPayload, updatedAt),
      };
    },

    async deletePage(targetPageId) {
      const sharedRecord = await fetchSharedRecord(client);
      const fallbackUpdatedAt = new Date().toISOString();
      const nextPayload = deleteWorkspacePage(sharedRecord.payload, targetPageId);
      const updatedAt = await upsertSharedPayload(client, nextPayload, fallbackUpdatedAt);
      return {
        updatedAt,
        pages: listPages(nextPayload, updatedAt),
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

async function upsertSharedPayload(client, payload, fallbackUpdatedAt) {
  const { data, error } = await client
    .from(SHARED_STATE_TABLE)
    .upsert({
      id: SHARED_STATE_ROW_ID,
      payload,
      updated_at: fallbackUpdatedAt,
    }, {
      onConflict: "id",
    })
    .select("updated_at")
    .single();

  if (error) {
    throw error;
  }

  return String(data?.updated_at || fallbackUpdatedAt);
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

function renameWorkspacePage(sharedPayload, pageId, nextTitle, updatedAt) {
  const normalizedPageId = normalizePageId(pageId);
  if (normalizedPageId === DEFAULT_PAGE_ID) {
    throw createWorkspaceError("workspace-primary-protected", "La page principale ne peut pas etre renommee.");
  }

  const container = normalizeWorkspaceContainer(sharedPayload);
  if (!hasWorkspacePage(container, normalizedPageId)) {
    throw createWorkspaceError("workspace-page-missing", "La page a renommer est introuvable.");
  }

  const existingMeta = normalizePageMeta(container.pageMeta?.[normalizedPageId], normalizedPageId);
  const title = normalizePageTitle(nextTitle) || existingMeta.title || defaultPageTitle(normalizedPageId);

  return {
    version: 3,
    pages: { ...container.pages },
    pageMeta: {
      ...container.pageMeta,
      [normalizedPageId]: {
        id: normalizedPageId,
        title,
        createdAt: existingMeta.createdAt || updatedAt,
        updatedAt,
      },
    },
  };
}

function deleteWorkspacePage(sharedPayload, pageId) {
  const normalizedPageId = normalizePageId(pageId);
  if (normalizedPageId === DEFAULT_PAGE_ID) {
    throw createWorkspaceError("workspace-primary-protected", "La page principale ne peut pas etre supprimee.");
  }

  const container = normalizeWorkspaceContainer(sharedPayload);
  if (!hasWorkspacePage(container, normalizedPageId)) {
    throw createWorkspaceError("workspace-page-missing", "La page a supprimer est introuvable.");
  }

  const nextPages = { ...container.pages };
  const nextPageMeta = { ...container.pageMeta };
  delete nextPages[normalizedPageId];
  delete nextPageMeta[normalizedPageId];

  return {
    version: 3,
    pages: nextPages,
    pageMeta: nextPageMeta,
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

function hasWorkspacePage(container, pageId) {
  return Object.prototype.hasOwnProperty.call(container.pages, pageId)
    || Object.prototype.hasOwnProperty.call(container.pageMeta, pageId);
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

function createWorkspaceError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}
