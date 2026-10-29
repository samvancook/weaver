import copy
import unittest
from weaver_runtime_sync import save_footage_event, get_footage_events, FIRESTORE_HANDOFF_ACTIONS
from weaver_runtime_db import FirestoreLedgerClient

class Ledger:
    def __init__(self):
        self.records = {}
        self.version = 0
        self.writes = 0
        self.conflict_once = False
    def get_raw_document_versioned(self, collection, key):
        return self.get_raw_document(collection,key), str(self.version)
    def get_raw_document(self, collection, key):
        return copy.deepcopy(self.records.get(key))
    def create_raw_document_if_absent(self, collection, key, record):
        if key in self.records: return None
        return self.save(key, record)
    def write_raw_document_if_unchanged(self, collection,key,record,version):
        if self.conflict_once:
            self.conflict_once=False; self.version+=1
            return None
        if version != str(self.version): return None
        return self.save(key,record)
    def save(self,key,record):
        self.writes+=1; self.version+=1; self.records[key]=copy.deepcopy(record)
        return record
    def list_raw_documents(self, collection):
        return copy.deepcopy(list(self.records.values()))

class FootageLedgerTest(unittest.TestCase):
    def record(self, digest='same'):
        return {'eventId':'inventory-123', 'digest':digest,'acceptedCount':2,'blockedCount':0}
    def test_saved_receipts_are_reloaded_and_stable_retries_do_not_duplicate_events(self):
        db=Ledger()
        first=save_footage_event(db,{'record':self.record()})
        second=save_footage_event(db,{'record':self.record()})
        self.assertTrue(first['ok']); self.assertTrue(second['ok'])
        self.assertEqual(first['record']['createdAt'],second['record']['createdAt'])
        self.assertEqual(len(get_footage_events(db,{})['records']),1)
        self.assertEqual(get_footage_events(db,{'eventId':'missing'})['records'],[])
    def test_changed_manifest_cannot_overwrite_existing_queue(self):
        db=Ledger(); save_footage_event(db,{'record':self.record()})
        original=copy.deepcopy(db.records); writes=db.writes
        result=save_footage_event(db,{'record':self.record('changed')})
        self.assertTrue(result['conflict']); self.assertEqual(db.records,original); self.assertEqual(db.writes,writes)
    def test_same_manifest_revalidation_retries_version_conflict(self):
        db=Ledger(); save_footage_event(db,{'record':self.record()}); db.conflict_once=True
        result=save_footage_event(db,{'record':{**self.record(),'blockedCount':1}})
        self.assertTrue(result['ok']); self.assertEqual(result['record']['blockedCount'],1)
    def test_sqlite_fallback_is_rejected(self):
        self.assertIn('save_footage_event',FIRESTORE_HANDOFF_ACTIONS)
        self.assertIn('get_footage_events',FIRESTORE_HANDOFF_ACTIONS)
        with self.assertRaises(ValueError): save_footage_event(object(),{'record':self.record()})
    def test_firestore_create_has_atomic_absent_precondition(self):
        client=object.__new__(FirestoreLedgerClient)
        client.document_url_for=lambda c,k:'https://example.invalid/doc'
        calls=[]
        client.request=lambda method,url,body:calls.append((method,url,body)) or {'fields':{}}
        client.create_raw_document_if_absent('events','id',self.record())
        self.assertEqual(calls[0][0],'PATCH')
        self.assertTrue(calls[0][1].endswith('?currentDocument.exists=false'))

if __name__=='__main__': unittest.main()
