"""Durable, at-most-once automatic attempts for newly approved video excerpts."""
import hashlib
import json
import re
import time
import unicodedata
import uuid
from datetime import datetime, timezone

COLLECTION = "excerptHandoffs"

def clean(value):
    return str(value or "").strip()

def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def file_id(record):
    payload = record.get("payload") or {}
    direct = clean(record.get("sourceVideoFileId") or payload.get("sourceVideoFileId"))
    match = re.search(r"/file/d/([\w-]+)", clean(record.get("sourceVideoUrl") or payload.get("sourceVideoUrl")))
    return direct or (match.group(1) if match else "")

def fingerprint(record):
    norm = lambda v: " ".join(unicodedata.normalize("NFKC", clean(v)).split()).lower()
    payload = record.get("payload") or {}
    fields = [clean(record.get("sourceRecordId")), norm(record.get("author")), norm(record.get("poemTitle")),
              norm(record.get("excerpt")), file_id(record), clean(record.get("sourceEvent") or payload.get("sourceEvent")),
              clean(record.get("bookTitle")), clean(record.get("bookShortener")), clean(record.get("releaseCatalog"))]
    return hashlib.sha256(json.dumps(fields, ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()

def key(record_id):
    return hashlib.sha256(record_id.encode()).hexdigest()[:32]

def write(db, record, version):
    return (db.write_raw_document_if_unchanged(COLLECTION, key(record["recordId"]), record, version)
            if version else db.create_raw_document_if_absent(COLLECTION, key(record["recordId"]), record))

def queue(db, payload):
    incoming = payload.get("handoffs") or []
    if not isinstance(incoming, list) or len(incoming) > 100:
        raise ValueError("Queue up to 100 exact video excerpts")
    for record in incoming:
        if (record.get("contentType") != "EXC" or record.get("recordId") != "weaver-exc-" + clean(record.get("sourceRecordId"))
                or not clean(record.get("author")) or not clean(record.get("excerpt"))
                or not re.fullmatch(r"[A-Za-z0-9_-]{15,100}", file_id(record))):
            raise ValueError("Exact approved video excerpt identity is required")
    if len({r["recordId"] for r in incoming}) != len(incoming):
        raise ValueError("Duplicate excerpt queue IDs")
    records, conflicts = [], []
    for source in incoming:
        old, version = db.get_raw_document_versioned(COLLECTION, key(source["recordId"]))
        current_fp, stamp = fingerprint(source), now_iso()
        if old:
            prior = old.get("delivery") or {}
            if fingerprint(old) == current_fp:
                # A repeated save cannot clear a receipt, lease, attempt or ambiguity.
                if prior.get("schemaVersion") == "1" or (old.get("handoffStatus") == "sent" and old.get("poetryPleaseItemId")):
                    if prior.get("state") == "needs_source_review" and prior.get("attempts") == 0 and not old.get("poetryPleaseItemId"):
                        record = {**old, **source, "handoffStatus":"queued", "delivery":{**prior,"state":"queued","queuedAt":stamp}}
                        saved = write(db, record, version)
                        if saved: records.append(saved)
                        else: conflicts.append(source["recordId"])
                    else:
                        records.append(old)
                    continue
                record = {**old, "delivery": {"schemaVersion": "1", "state": "needs_reconciliation", "sourceFingerprint": current_fp,
                          "attempts": 1, "queuedAt": stamp, "checkedAt": stamp}, "handoffStatus": "failed",
                          "errorMessage": "Existing delivery history requires receipt reconciliation; no automatic resend"}
            elif prior.get("schemaVersion") == "1" and prior.get("attempts") == 0 and prior.get("state") != "sending" and not old.get("poetryPleaseItemId"):
                record = {**old, **source, "handoffStatus":"queued", "delivery":{**prior,"state":"queued","sourceFingerprint":current_fp,"queuedAt":stamp},
                          "deliveryHistory":(list(old.get("deliveryHistory") or []) + [{"at":stamp,"event":"unattempted_source_updated","previousFingerprint":fingerprint(old)}])[-20:]}
            else:
                record = {**old, "delivery": {**prior, "schemaVersion": "1", "state": "needs_source_review",
                          "sourceFingerprint": fingerprint(old), "pendingSourceFingerprint": current_fp, "checkedAt": stamp},
                          "handoffStatus": "failed", "errorMessage": "Current approval differs from the saved delivery version"}
        else:
            record = {**source, "handoffStatus": "queued", "handoffMode": "scheduled_video_excerpt",
                      "delivery": {"schemaVersion": "1", "state": "queued", "sourceFingerprint": current_fp,
                                   "attempts": 0, "queuedAt": stamp}, "createdAt": stamp}
        record["ledgerUpdatedAt"] = stamp
        saved = write(db, record, version)
        if saved:
            records.append(saved)
        else:
            conflicts.append(source["recordId"])
    return {"ok": not conflicts, "records": records, "conflicts": conflicts, "savedCount": len(records),
            "needsReviewCount": sum((r.get("delivery") or {}).get("state") in {"needs_source_review", "needs_reconciliation"} for r in records)}

def get_queue(db, payload):
    return {"ok": True, "records": [r for r in db.list_raw_documents(COLLECTION)
            if (r.get("delivery") or {}).get("schemaVersion") == "1"]}

def claim(db, payload):
    expected = payload.get("expectedRecord") or {}
    old, version = db.get_raw_document_versioned(COLLECTION, key(clean(expected.get("recordId"))))
    delivery = (old or {}).get("delivery") or {}
    if (old != expected or delivery.get("schemaVersion") != "1" or delivery.get("state") not in {"queued", "waiting_video"}
            or delivery.get("attempts") != 0 or payload.get("sourceFingerprint") != delivery.get("sourceFingerprint")
            or fingerprint(old) != delivery.get("sourceFingerprint")):
        return {"ok": True, "claimed": False, "reason": "delivery_changed_or_already_attempted"}
    token, stamp = uuid.uuid4().hex, now_iso()
    record = {**old, "delivery": {**delivery, "state": "sending", "attempts": 1, "token": token,
              "leaseExpiresAt": time.time() + 120, "startedAt": stamp, "checkedAt": stamp}, "ledgerUpdatedAt": stamp}
    saved = write(db, record, version)
    return {"ok": True, "claimed": bool(saved), "record": saved, "token": token if saved else ""}

def finish(db, payload):
    record_id = clean(payload.get("recordId"))
    old, version = db.get_raw_document_versioned(COLLECTION, key(record_id))
    delivery = (old or {}).get("delivery") or {}
    if (delivery.get("schemaVersion") != "1" or delivery.get("state") != "sending"
            or not payload.get("token") or payload["token"] != delivery.get("token")
            or float(delivery.get("leaseExpiresAt") or 0) <= time.time()):
        return {"ok": False, "saved": False, "reason": "delivery_lease_changed"}
    accepted = payload.get("acknowledged") is True and bool(clean(payload.get("canonicalExcerptId")))
    canonical = clean(payload.get("canonicalExcerptId")) if accepted else clean(old.get("poetryPleaseItemId"))
    if old.get("poetryPleaseItemId") and old["poetryPleaseItemId"] != canonical:
        return {"ok": False, "saved": False, "reason": "receiver_identity_changed"}
    stamp = now_iso()
    record = {**old, "handoffStatus": "sent" if accepted else "failed", "poetryPleaseItemId": canonical,
              "handedOffAt": stamp if accepted else old.get("handedOffAt", ""),
              "errorMessage": "" if accepted else clean(payload.get("error"))[:500],
              "delivery": {**delivery, "state": "acknowledged" if accepted else "needs_reconciliation",
                           "leaseExpiresAt": 0, "checkedAt": stamp}, "ledgerUpdatedAt": stamp}
    saved = write(db, record, version)
    return {"ok": bool(saved), "saved": bool(saved), "record": saved}

def review(db, payload):
    expected = payload.get("expectedRecord") or {}
    old, version = db.get_raw_document_versioned(COLLECTION, key(clean(expected.get("recordId"))))
    delivery = (old or {}).get("delivery") or {}
    if old != expected or delivery.get("schemaVersion") != "1":
        return {"ok": False, "saved": False, "reason": "delivery_changed"}
    reason = payload.get("reason")
    states = {"waiting_video": "waiting_video", "source_changed": "needs_source_review", "approval_withdrawn": "needs_source_review",
              "expired": "needs_reconciliation", "inspection_failed": delivery.get("state")}
    if reason not in states:
        raise ValueError("Unsupported excerpt delivery transition")
    if reason == "expired" and (delivery.get("state") != "sending" or float(delivery.get("leaseExpiresAt") or 0) > time.time()):
        return {"ok": False, "saved": False, "reason": "delivery_still_active"}
    if reason != "expired" and delivery.get("state") == "sending":
        return {"ok": False, "saved": False, "reason": "delivery_still_active"}
    stamp = now_iso()
    error = clean(payload.get("error"))[:500]
    record = {**old, "delivery": {**delivery, "state": states[reason], "checkedAt": stamp, "lastCheckError": error}, "ledgerUpdatedAt": stamp}
    if reason in {"source_changed", "approval_withdrawn", "expired"}:
        record.update(handoffStatus="failed", errorMessage=error or "Delivery needs reconciliation; no automatic resend")
    saved = write(db, record, version)
    return {"ok": bool(saved), "saved": bool(saved), "record": saved}
