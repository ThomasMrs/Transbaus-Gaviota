const SUPABASE_URL = "https://znbcnahjvtdndttqjpec.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_L7q8_g_9SavtGfkNlPjK6Q_imFwj6_N";
const SHARED_STATE_TABLE = "shared_state";
const SHARED_STATE_ROW_ID = "global";

export function createSharedStateStore() {
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
      const { data, error } = await client
        .from(SHARED_STATE_TABLE)
        .select("payload, updated_at")
        .eq("id", SHARED_STATE_ROW_ID)
        .maybeSingle();

      if (error) {
        throw error;
      }

      return {
        state: data?.payload || null,
        updatedAt: String(data?.updated_at || ""),
      };
    },

    async saveStateRecord(payload) {
      const fallbackUpdatedAt = new Date().toISOString();
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

      return {
        updatedAt: String(data?.updated_at || fallbackUpdatedAt),
      };
    },
  };
}
