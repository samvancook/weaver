import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Verified unique matches from button-poetry-catalog's 2026-10-15 excerpt seed.
export const POEM_KEY_SEED_SOURCE = "catalog_shared_poem_key_seed_2026-10-15";
const seedPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "data", "catalog_poem_key_seed_2026-10-15.json");
const keysByExcerptHash = JSON.parse(readFileSync(seedPath, "utf8"));

export function fingerprintExcerptForCatalogKey(value) {
  const normalized = String(value || "")
    .replace(/\*([^*\n]+)\*/g, "$1")
    .replace(/\*/g, "")
    .toLowerCase()
    .replace(/[—–]/g, " ")
    .replace(/&/g, " and ")
    .replace(/[’‘]/g, "'")
    .replace(/["“”`]/g, "")
    .replace(/(?<![a-z0-9])'|'(?![a-z0-9])/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[^a-z0-9\s']/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return createHash("sha256").update(normalized).digest("hex");
}

export function poemKeyForExcerptText(value) {
  if (!String(value || "").trim()) return null;
  return keysByExcerptHash[fingerprintExcerptForCatalogKey(value)] || null;
}
