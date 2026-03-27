import http from "node:http";
import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT || 4173);

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
    const child = spawn("python3", [path.join(__dirname, "catalog_validate.py")]);
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
    res.writeHead(200, {
      "Content-Type": contentTypes[ext] || "application/octet-stream"
    });
    res.end(data);
  } catch {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  }
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host}`);

  if (url.pathname === "/api/bootstrap") {
    return sendJson(res, 200, {
      appName: "Excerpt Review Tool",
      sections: [
        { id: "intake", label: "Intake Queue" },
        { id: "duplicates", label: "Duplicate Review" },
        { id: "approved", label: "Approved" },
        { id: "exports", label: "Export Queue" }
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
      return sendJson(res, 500, { ok: false, error: error.message });
    }
  }

  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const safePath = path.normalize(requestedPath).replace(/^(\.\.[/\\])+/, "");
  const filePath = path.join(publicDir, safePath);
  return serveFile(res, filePath);
});

server.listen(port, () => {
  console.log(`Weaver dev server running at http://127.0.0.1:${port}`);
});
