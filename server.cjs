const fs = require("node:fs");
const fsp = require("node:fs/promises");
const http = require("node:http");
const path = require("node:path");
const { DatabaseSync } = require("node:sqlite");
const { URL } = require("node:url");

const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number.parseInt(process.env.PORT || "4173", 10);
const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, "data");
const DB_PATH = path.join(DATA_DIR, "transbaus.sqlite");
const JSON_LIMIT_BYTES = 2 * 1024 * 1024;
const STATE_ROW_ID = "shared-state";
const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".jpg": "image/jpeg",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".webp": "image/webp",
};

fs.mkdirSync(DATA_DIR, { recursive: true });

const db = new DatabaseSync(DB_PATH);
db.exec(`
  PRAGMA journal_mode = WAL;
  CREATE TABLE IF NOT EXISTS app_state (
    id TEXT PRIMARY KEY,
    payload TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );
`);

const selectStateStatement = db.prepare(`
  SELECT payload, updated_at
  FROM app_state
  WHERE id = ?
`);
const upsertStateStatement = db.prepare(`
  INSERT INTO app_state (id, payload, updated_at)
  VALUES (?, ?, ?)
  ON CONFLICT(id) DO UPDATE SET
    payload = excluded.payload,
    updated_at = excluded.updated_at
`);

const server = http.createServer(async (request, response) => {
  const method = String(request.method || "GET").toUpperCase();
  const requestUrl = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`);

  if (requestUrl.pathname === "/api/health") {
    if (method !== "GET") {
      sendJson(response, 405, { error: "method-not-allowed" });
      return;
    }

    sendJson(response, 200, {
      ok: true,
      databasePath: DB_PATH,
      updatedAt: readSharedStateRecord().updatedAt,
    });
    return;
  }

  if (requestUrl.pathname === "/api/state") {
    if (method === "GET") {
      sendJson(response, 200, readSharedStateRecord());
      return;
    }

    if (method === "PUT") {
      try {
        const payload = await parseJsonBody(request);
        const validationError = validateSharedStatePayload(payload);
        if (validationError) {
          sendJson(response, 400, { error: validationError });
          return;
        }

        sendJson(response, 200, writeSharedStateRecord(payload));
      } catch (error) {
        if (error?.message === "invalid-json") {
          sendJson(response, 400, { error: "invalid-json" });
          return;
        }

        if (error?.message === "payload-too-large") {
          sendJson(response, 413, { error: "payload-too-large" });
          return;
        }

        console.error("Erreur API /api/state", error);
        sendJson(response, 500, { error: "server-error" });
      }
      return;
    }

    sendJson(response, 405, { error: "method-not-allowed" });
    return;
  }

  if (!["GET", "HEAD"].includes(method)) {
    sendJson(response, 405, { error: "method-not-allowed" });
    return;
  }

  await serveStaticFile(requestUrl.pathname, method, response);
});

server.listen(PORT, HOST, () => {
  const hostLabel = HOST === "0.0.0.0" ? "localhost" : HOST;
  console.log(`Serveur partage actif sur http://${hostLabel}:${PORT}`);
  console.log(`BDD SQLite : ${DB_PATH}`);
});

process.on("SIGINT", () => closeServer("SIGINT"));
process.on("SIGTERM", () => closeServer("SIGTERM"));

function readSharedStateRecord() {
  const row = selectStateStatement.get(STATE_ROW_ID);
  if (!row) {
    return {
      state: null,
      updatedAt: "",
    };
  }

  try {
    return {
      state: JSON.parse(String(row.payload || "null")),
      updatedAt: String(row.updated_at || ""),
    };
  } catch (error) {
    console.error("Etat partage illisible dans SQLite", error);
    return {
      state: null,
      updatedAt: String(row.updated_at || ""),
    };
  }
}

function writeSharedStateRecord(nextState) {
  const updatedAt = new Date().toISOString();
  upsertStateStatement.run(STATE_ROW_ID, JSON.stringify(nextState), updatedAt);
  return {
    state: nextState,
    updatedAt,
  };
}

function validateSharedStatePayload(payload) {
  if (!payload || typeof payload !== "object") {
    return "missing-json-body";
  }

  const requiredArrays = ["baques", "parcels", "deliveryNotes", "destinationRules"];
  for (const key of requiredArrays) {
    if (!Array.isArray(payload[key])) {
      return `invalid-${key}`;
    }
  }

  return "";
}

function parseJsonBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalBytes = 0;
    let settled = false;

    request.on("data", (chunk) => {
      if (settled) {
        return;
      }

      totalBytes += chunk.length;
      if (totalBytes > JSON_LIMIT_BYTES) {
        settled = true;
        reject(new Error("payload-too-large"));
        return;
      }

      chunks.push(Buffer.from(chunk));
    });

    request.on("end", () => {
      if (settled) {
        return;
      }

      settled = true;
      const rawBody = Buffer.concat(chunks).toString("utf8").trim();
      if (!rawBody) {
        resolve(null);
        return;
      }

      try {
        resolve(JSON.parse(rawBody));
      } catch (error) {
        reject(new Error("invalid-json"));
      }
    });

    request.on("error", (error) => {
      if (settled) {
        return;
      }

      settled = true;
      reject(error);
    });
  });
}

async function serveStaticFile(pathname, method, response) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  let decodedPath;

  try {
    decodedPath = decodeURIComponent(requestedPath);
  } catch (error) {
    sendText(response, 400, "bad-request");
    return;
  }

  const absolutePath = path.resolve(ROOT_DIR, `.${decodedPath}`);
  const relativePath = path.relative(ROOT_DIR, absolutePath);

  if (relativePath.startsWith("..") || path.isAbsolute(relativePath)) {
    sendText(response, 403, "forbidden");
    return;
  }

  try {
    const stats = await fsp.stat(absolutePath);
    if (stats.isDirectory()) {
      await serveStaticFile(`${decodedPath.replace(/\/+$/, "")}/index.html`, method, response);
      return;
    }

    const extension = path.extname(absolutePath).toLowerCase();
    response.writeHead(200, {
      "Cache-Control": extension === ".html" ? "no-cache" : "public, max-age=300",
      "Content-Type": MIME_TYPES[extension] || "application/octet-stream",
      "Content-Length": stats.size,
    });

    if (method === "HEAD") {
      response.end();
      return;
    }

    fs.createReadStream(absolutePath).pipe(response);
  } catch (error) {
    if (error?.code === "ENOENT") {
      if (!path.extname(decodedPath)) {
        await serveStaticFile("/index.html", method, response);
        return;
      }

      sendText(response, 404, "not-found");
      return;
    }

    console.error("Erreur fichier statique", error);
    sendText(response, 500, "server-error");
  }
}

function sendJson(response, statusCode, payload) {
  const body = JSON.stringify(payload);
  response.writeHead(statusCode, {
    "Cache-Control": "no-store",
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
  });
  response.end(body);
}

function sendText(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Cache-Control": "no-store",
    "Content-Type": "text/plain; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
  });
  response.end(payload);
}

function closeServer(signal) {
  console.log(`Arret du serveur (${signal})`);
  server.close(() => {
    db.close();
    process.exit(0);
  });
}
