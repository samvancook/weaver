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

function escapeHtml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function getLibraryStatusLabel(status) {
  if (!status) {
    return "";
  }
  if (status.made) {
    return "Graphic appears to have already been made.";
  }
  if (status.approvedForQi) {
    return "Approved for quote image, but not confirmed made.";
  }
  return "In excerpt library only; no quote-image status found.";
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

async function proxySaveReviews(apiBaseUrl, updates) {
  if (!apiBaseUrl) {
    throw new Error("Missing Apps Script Web App URL.");
  }

  const url = new URL(apiBaseUrl);
  url.searchParams.set("action", "saveReviews");

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ updates })
  });

  const text = await response.text();
  let data;
  try {
    data = JSON.parse(text || "{}");
  } catch (_error) {
    throw new Error("Apps Script batch save returned invalid JSON.");
  }

  if (!response.ok || !data.ok) {
    throw new Error(data.error || `Apps Script batch save failed with status ${response.status}.`);
  }

  return data;
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

function runCatalogPoemLookup(bookTitle, poemTitle, excerptText) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "catalog_poem_text.py")], {
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
        reject(new Error(stderr || `catalog_poem_text.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ bookTitle, poemTitle, excerptText }));
    child.stdin.end();
  });
}

function runLibraryExcerptLookup(sourceRow) {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", [path.join(__dirname, "excerpt_library_text.py")], {
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
        reject(new Error(stderr || `excerpt_library_text.py exited with code ${code}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(error);
      }
    });

    child.stdin.write(JSON.stringify({ sourceRow }));
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

  if (url.pathname === "/api/save-reviews" && req.method === "POST") {
    try {
      const body = await readRequestBody(req);
      const parsed = JSON.parse(body || "{}");
      const apiBaseUrl = typeof parsed.apiBaseUrl === "string" ? parsed.apiBaseUrl.trim() : "";
      const updates = Array.isArray(parsed.updates) ? parsed.updates : [];

      const result = await proxySaveReviews(apiBaseUrl, updates);
      return sendJson(res, 200, result);
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  if (url.pathname === "/catalog-poem" && req.method === "GET") {
    try {
      const bookTitle = url.searchParams.get("bookTitle") || "";
      const poemTitle = url.searchParams.get("poemTitle") || "";
      const excerptText = url.searchParams.get("excerptText") || "";
      const result = await runCatalogPoemLookup(bookTitle, poemTitle, excerptText);
      const html = result.ok
        ? `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(result.poemTitle)} · Weaver Catalog Context</title>
  <style>
    body { margin: 0; font-family: Georgia, "Times New Roman", serif; background: #f7f2ea; color: #1d1a17; }
    main { max-width: 820px; margin: 0 auto; padding: 32px 24px 56px; }
    .meta { color: #6b6259; font-size: 14px; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 12px; }
    h1 { margin: 0 0 8px; font-size: clamp(30px, 5vw, 46px); line-height: 1.05; }
    h2 { margin: 0 0 22px; font-size: 22px; color: #b84f2d; font-weight: 600; }
    pre { white-space: pre-wrap; word-break: break-word; background: rgba(255,252,247,.92); border: 1px solid rgba(29,26,23,.1); border-radius: 24px; padding: 24px; font: 18px/1.7 Georgia, "Times New Roman", serif; box-shadow: 0 18px 40px rgba(62,39,27,.08); }
  </style>
</head>
<body>
  <main>
    <div class="meta">Catalog Poem Context · ${escapeHtml(result.wordCount || "")} words</div>
    <h1>${escapeHtml(result.poemTitle)}</h1>
    <h2>${escapeHtml(result.author)} · ${escapeHtml(result.bookTitle)}</h2>
    <pre>${escapeHtml(result.text)}</pre>
  </main>
</body>
</html>`
        : `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" /><title>Weaver Catalog Context</title></head>
<body style="font-family: Georgia, 'Times New Roman', serif; background:#f7f2ea; color:#1d1a17; padding:32px;">
  <h1 style="margin-top:0;">Catalog poem unavailable</h1>
  <p>${escapeHtml(result.error || "Unable to load poem text.")}</p>
</body>
</html>`;
      res.writeHead(result.ok ? 200 : 404, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
      return;
    } catch (error) {
      res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
      res.end(`<html><body style="font-family: sans-serif; padding: 32px;"><h1>Catalog poem lookup failed</h1><p>${escapeHtml(error.message)}</p></body></html>`);
      return;
    }
  }

  if (url.pathname === "/library-excerpt" && req.method === "GET") {
    try {
      const sourceRow = Number(url.searchParams.get("sourceRow") || 0);
      const result = await runLibraryExcerptLookup(sourceRow);
      const libraryStatusLabel = getLibraryStatusLabel(result.libraryStatus);
      const html = result.ok
        ? `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Excerpt Library Match · ${escapeHtml(result.poemTitle || "Weaver")}</title>
  <style>
    body { margin: 0; font-family: Georgia, "Times New Roman", serif; background: #f7f2ea; color: #1d1a17; }
    main { max-width: 820px; margin: 0 auto; padding: 32px 24px 56px; }
    .meta { color: #6b6259; font-size: 14px; text-transform: uppercase; letter-spacing: .08em; margin-bottom: 12px; }
    .status { margin: 0 0 16px; color: #6b6259; font-size: 16px; }
    h1 { margin: 0 0 8px; font-size: clamp(30px, 5vw, 46px); line-height: 1.05; }
    h2 { margin: 0 0 22px; font-size: 22px; color: #b84f2d; font-weight: 600; }
    pre { white-space: pre-wrap; word-break: break-word; background: rgba(255,252,247,.92); border: 1px solid rgba(29,26,23,.1); border-radius: 24px; padding: 24px; font: 18px/1.7 Georgia, "Times New Roman", serif; box-shadow: 0 18px 40px rgba(62,39,27,.08); }
  </style>
</head>
<body>
  <main>
    <div class="meta">Excerpt Library Match · Row ${escapeHtml(result.sourceRow)}${result.wordCount ? ` · ${escapeHtml(result.wordCount)} words` : ""}</div>
    <h1>${escapeHtml(result.poemTitle || "Untitled")}</h1>
    <h2>${escapeHtml(result.author || "Unknown author")}${result.bookTitle ? ` · ${escapeHtml(result.bookTitle)}` : ""}</h2>
    ${libraryStatusLabel ? `<p class="status">${escapeHtml(libraryStatusLabel)}</p>` : ""}
    <pre>${escapeHtml(result.text)}</pre>
  </main>
</body>
</html>`
        : `<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" /><title>Excerpt Library Match</title></head>
<body style="font-family: Georgia, 'Times New Roman', serif; background:#f7f2ea; color:#1d1a17; padding:32px;">
  <h1 style="margin-top:0;">Library excerpt unavailable</h1>
  <p>${escapeHtml(result.error || "Unable to load excerpt text.")}</p>
</body>
</html>`;
      res.writeHead(result.ok ? 200 : 404, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
      return;
    } catch (error) {
      res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
      res.end(`<html><body style="font-family: sans-serif; padding: 32px;"><h1>Library excerpt lookup failed</h1><p>${escapeHtml(error.message)}</p></body></html>`);
      return;
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
