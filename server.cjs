const fs = require("node:fs");
const fsp = require("node:fs/promises");
const http = require("node:http");
const path = require("node:path");
const { URL } = require("node:url");

const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number.parseInt(process.env.PORT || "4173", 10);
const ROOT_DIR = __dirname;
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

const server = http.createServer(async (request, response) => {
  const method = String(request.method || "GET").toUpperCase();
  const requestUrl = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`);

  if (!["GET", "HEAD"].includes(method)) {
    sendText(response, 405, "method-not-allowed");
    return;
  }

  await serveStaticFile(requestUrl.pathname, method, response);
});

server.listen(PORT, HOST, () => {
  const hostLabel = HOST === "0.0.0.0" ? "localhost" : HOST;
  console.log(`Serveur local actif sur http://${hostLabel}:${PORT}`);
  console.log("Les donnees metier sont synchronisees directement avec Supabase.");
});

process.on("SIGINT", () => closeServer("SIGINT"));
process.on("SIGTERM", () => closeServer("SIGTERM"));

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
    process.exit(0);
  });
}
