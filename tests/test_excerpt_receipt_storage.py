import copy
import time
import sqlite3
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch
from weaver_runtime_sync import (fetch_excerpt_handoffs, sync_excerpt_handoffs, reconcile_excerpt_handoff,
    excerpt_receipt_fingerprint, excerpt_handoff_document_id, FIRESTORE_HANDOFF_ACTIONS)
from weaver_runtime_sync import legacy_excerpt_handoffs
from weaver_runtime_db import connect_runtime_db, ensure_runtime_schema, upsert_excerpt_handoff


class Store:
    def __init__(self, rows=None, conflict=False):
        self.rows=rows if rows is not None else {}
        self.conflict=conflict
        self.writes=0
    def get_raw_document(self, collection, key): return copy.deepcopy(self.rows.get(key))
    def get_raw_document_versioned(self, collection, key):
        return self.get_raw_document(collection,key), 'version' if key in self.rows else ''
    def list_raw_documents(self, collection): return copy.deepcopy(list(self.rows.values()))
    def create_raw_document_if_absent(self, collection, key, row):
        if self.conflict or key in self.rows: return None
        self.rows[key]=copy.deepcopy(row);self.writes+=1;return copy.deepcopy(row)
    def write_raw_document_if_unchanged(self, collection,key,row,version):
        if self.conflict or version!='version': return None
        self.rows[key]=copy.deepcopy(row);self.writes+=1;return copy.deepcopy(row)


class ExcerptReceiptStorageTest(unittest.TestCase):
    def setUp(self):
        self.record={'recordId':'handoff','sourceRecordId':'source','contentType':'EXC','excerpt':'A line\nwith meaning.',
            'handoffStatus':'failed','handedOffAt':'old-time','errorMessage':'old 502','poetryPleaseItemId':'',
            'approvedAt':'original approval','payload':{'sourceVideoFileId':'original','sourceRow':123}}
        self.receipt={'verified':True,'sourceRecordId':'source','canonicalExcerptId':'EXC-ONE',
            'excerptFingerprint':excerpt_receipt_fingerprint(self.record['excerpt']),
            'checkedAtMs':time.time()*1000,'importLedgerIds':['exact-ledger']}
        self.payload={'expectedRecord':self.record,'receipt':self.receipt,'checkedBy':'operator'}

    def test_fresh_container_without_legacy_database_uses_durable_records_without_creating_sqlite(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'absent.db'
            store=Store({excerpt_handoff_document_id('handoff'):self.record})
            with patch('weaver_runtime_sync.DEFAULT_DB_PATH',path):
                self.assertEqual(fetch_excerpt_handoffs(store,{})['records'],[self.record])
                self.assertFalse(path.exists())
                self.assertFalse(reconcile_excerpt_handoff(Store(),self.payload)['saved'])
                result=reconcile_excerpt_handoff(Store(),{**self.payload,'approvedBaseline':True})
                self.assertTrue(result['saved'])
                self.assertEqual(result['record']['reconciliation']['previousError'],'old 502')
                self.assertFalse(path.exists())

    def test_existing_legacy_database_is_read_only_and_corruption_is_not_hidden(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'legacy.db'
            with patch('weaver_runtime_sync.DEFAULT_DB_PATH',path):
                sqlite3.connect(path).close()
                self.assertEqual(legacy_excerpt_handoffs([]),[])
                connection=connect_runtime_db(path)
                ensure_runtime_schema(connection)
                upsert_excerpt_handoff(connection,self.record)
                connection.commit();connection.close()
                self.assertEqual(legacy_excerpt_handoffs(['handoff'])[0]['sourceRecordId'],'source')
                path.write_bytes(b'corrupt database')
                with self.assertRaises(sqlite3.DatabaseError): legacy_excerpt_handoffs([])

    def test_approved_baseline_cannot_override_an_existing_durable_row(self):
        store=Store({excerpt_handoff_document_id('handoff'):{**self.record,'excerpt':'Changed'}})
        self.assertFalse(reconcile_excerpt_handoff(store,{**self.payload,'approvedBaseline':True})['saved'])
        self.assertEqual(store.writes,0)

    def test_new_status_survives_another_instance_and_overrides_legacy(self):
        store=Store()
        self.assertTrue({'upsert_excerpt_handoffs','get_excerpt_handoffs','reconcile_excerpt_handoff'} <= FIRESTORE_HANDOFF_ACTIONS)
        sync_excerpt_handoffs(store,{'handoffs':[self.record]})
        saved={**self.record,'handoffStatus':'sent','poetryPleaseItemId':'EXC-ONE'}
        sync_excerpt_handoffs(Store(store.rows),{'handoffs':[saved]})
        with patch('weaver_runtime_sync.legacy_excerpt_handoffs',return_value=[self.record]):
            result=fetch_excerpt_handoffs(Store(store.rows),{})
        self.assertEqual(len(result['records']),1)
        self.assertEqual(result['records'][0]['handoffStatus'],'sent')
        self.assertEqual(result['backend'],'firestore')

    def test_receipt_only_repair_preserves_content_approval_and_original_failure_audit(self):
        store=Store()
        with patch('weaver_runtime_sync.legacy_excerpt_handoffs',return_value=[self.record]):
            result=reconcile_excerpt_handoff(store,self.payload)
        self.assertTrue(result['saved']);saved=result['record']
        for key in ['excerpt','approvedAt','handedOffAt','payload','sourceRecordId']:
            self.assertEqual(saved[key],self.record[key])
        self.assertEqual(saved['reconciliation']['previousError'],'old 502')
        self.assertEqual(saved['reconciliation']['mode'],'receipt_only_no_resend')
        again=reconcile_excerpt_handoff(store,{**self.payload,'expectedRecord':saved})
        self.assertTrue(again['alreadyVerified']);self.assertEqual(store.writes,1)

    def test_stale_wrong_source_text_or_receiver_and_concurrent_changes_hold(self):
        for change in [{'verified':False},{'checkedAtMs':0},{'sourceRecordId':'other'},
                       {'excerptFingerprint':'different'},{'importLedgerIds':[]},{'canonicalExcerptId':''}]:
            store=Store({excerpt_handoff_document_id('handoff'):copy.deepcopy(self.record)})
            result=reconcile_excerpt_handoff(store,{**self.payload,'receipt':{**self.receipt,**change}})
            self.assertFalse(result['saved']);self.assertEqual(store.writes,0)
        store=Store({excerpt_handoff_document_id('handoff'):{**self.record,'excerpt':'Changed'}})
        self.assertFalse(reconcile_excerpt_handoff(store,self.payload)['saved'])
        with patch('weaver_runtime_sync.legacy_excerpt_handoffs',return_value=[self.record]):
            self.assertFalse(reconcile_excerpt_handoff(Store(conflict=True),self.payload)['saved'])

    def test_verified_receipt_cannot_be_overwritten_by_old_retry(self):
        store=Store({excerpt_handoff_document_id('handoff'):{**self.record,'handoffStatus':'sent',
            'poetryPleaseItemId':'EXC-ONE','reconciliation':{'previousError':'old 502'}}})
        with self.assertRaises(ValueError): sync_excerpt_handoffs(store,{'handoffs':[self.record]})
        self.assertEqual(store.writes,0)

if __name__=='__main__': unittest.main()
