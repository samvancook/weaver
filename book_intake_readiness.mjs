function key(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

export function buildPublishingBookReadiness(publishingBooks = [], catalogBooks = [], lookupStatus = "unknown") {
  const titles = new Set(catalogBooks.map(book => key(book.title)).filter(Boolean));
  const shorteners = new Set(catalogBooks.map(book => key(book.bookShortener)).filter(Boolean));
  return publishingBooks.map(book => ({
    title: String(book.title || "").trim(),
    bookShortener: String(book.bookShortener || "").trim(),
    releaseCatalog: String(book.releaseCatalog || "").trim(),
    status: lookupStatus !== "live" ? "unverified"
      : titles.has(key(book.title)) || (key(book.bookShortener) && shorteners.has(key(book.bookShortener)))
        ? "catalog_backed" : "metadata_only"
  })).filter(book => book.title);
}
