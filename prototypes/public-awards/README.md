# Button Poetry awards browser — review prototype

This is a **standalone public-facing awards UI prototype** stored inside the Weaver repository for engineering review. It is **not** a Weaver admin feature, **not** wired to the current production site, and **not** deployed. No public /awards route or production data sync is created by this branch.

## Preview

Serve the repository locally (any static HTTP server will do):

```bash
cd prototypes/public-awards
python3 -m http.server 8000
```

Open http://localhost:8000 . Do not open `index.html` directly as a `file://` URL; the browser blocks local JSON fetches. The interface contains search, year/result/organization filters, a clear action, shareable query-string filters, and book + organizer links. It works without a framework or build step.

## Current dataset

`awards.json` contains six candidate public recognitions across four books, researched against award-organizer sources, with product permalinks and print ISBNs. It is explicitly `status: "review_only"` and **cannot be assumed approved for publication**. Detailed research:
[Codex Awards pilot brief](https://docs.google.com/document/d/1udmo56AbsoEKUQBV6Zhw4SBTmDmuaCf7s3OLs45198o/edit).

The six source IDs are stable slugs selected for this pilot. `sheet_row` is audit provenance only, and must never be used as a durable production key. All links must be verified again before launch.

## Check fixtures

Run `node tests/validate.mjs` from this directory with Node 20+. It rejects missing fields, duplicate recognition IDs, non-HTTPS URLs, duplicate book/award/year/result combinations and unapproved non-external categories.

## Integration roadmap

1. Identify the actual **public** buttonpoetry.com application/repository and deployment workflow. `samvancook/weaver` is an internal excerpt/publishing management app (Firebase Hosting/Cloud Run), not necessarily the public WooCommerce storefront.
2. Map records to catalog product IDs from the authoritative commerce API (current mappings are *title/author/ISBN/permalink*, not numeric WooCommerce IDs).
3. Add an editorial approval/withhold layer and a secure, read-only export from the Awards & Honors source workbook. Do **not** expose private Google Sheet contents, staff notes, submission data, account references, or sticker research.
4. Integrate the UI into the real public site's design system and URLs. Test mobile, keyboard, accessibility, security, and broken source links; implement error handling and rollback.
5. Obtain explicit approval before publishing, then expand from six pilot recognitions to the eligible, independently verified subset of the 51 external records.

## Source of truth

- [Awards & Honors spreadsheet](https://docs.google.com/spreadsheets/d/1Xk9Z-7cOU3dW76iwVLvy4IzwkTxyq5dWYV6Pjc/edit)
- [Awards roadmap](https://docs.google.com/document/d/1wbufSeYQERaYZLmW7hCDSIXFQqce51Q8HJ3Xo99SbPY/edit)
- [Classification/publishing rules](https://docs.google.com/document/d/1yFbMl-24NvHMDumknyikSgMobH8jwMdPLHfx3ccZuxY/edit)
