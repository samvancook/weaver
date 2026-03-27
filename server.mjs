import http from "node:http";
import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT || 8080);

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml"
};

function sendJson(res, statusCode, body) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body, null, 2));
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => {
      body += chunk.toString();
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

function runCatalogValidation(records) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "catalog_validate.py")], {
      env: {
        ...process.env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", chunk => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", chunk => {
      stderr += chunk.toString();
    });
    child.on("error", reject);
    child.on("close", code => {
      if (code !== 0) {
        reject(new Error(stderr || `catalog_validate.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ records }));
    child.stdin.end();
  });
}

async function serveFile(res, filePath) {
  try {
    const data = await readFile(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const headers = {
      "Content-Type": contentTypes[ext] || "application/octet-stream"
    };

    if (ext === ".html") {
      headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0";
    } else if (ext === ".js" || ext === ".css") {
      headers["Cache-Control"] = "public, max-age=31536000, immutable";
    }

    res.writeHead(200, headers);
    res.end(data);
  } catch {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host}`);

  if (url.pathname === "/health") {
    return sendJson(res, 200, { ok: true, service: "weaver-web" });
  }

  if (url.pathname === "/api/bootstrap") {
    return sendJson(res, 200, {
      appName: "Weaver",
      sections: [
        { id: "review", label: "Review queue" },
        { id: "corrections", label: "Needs correction" }
      ]
    });
  }

  if (url.pathname === "/api/catalog/validate" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const records = Array.isArray(parsed.records) ? parsed.records : [];
      const result = await runCatalogValidation(records);
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const safePath = path.normalize(requestedPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(publicDir, safePath);

  if (!filePath.startsWith(publicDir)) {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Forbidden");
    return;
  }

  return serveFile(res, filePath);
});

server.listen(port, () => {
  console.log(`Weaver server running on port ${port}`);
});
