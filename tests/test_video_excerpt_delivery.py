import copy
import time
import unittest
from test_video_pipeline import Store
from video_excerpt_delivery import queue, claim, finish, review, fingerprint, key, COLLECTION

class VideoExcerptQueueTests(unittest.TestCase):
    def record(self):
        return {'recordId':'weaver-exc-source','sourceRecordId':'source','contentType':'EXC','author':'Poet','poemTitle':'Poem','excerpt':'Approved lines',
                'payload':{'sourceVideoFileId':'video_file_123456789','sourceEvent':'Event'},'handoffStatus':'queued'}
    def queued(self,db):return queue(db,{'handoffs':[self.record()]})['records'][0]
    def test_repeat_approval_preserves_one_queue_and_active_attempt(self):
        db=Store();r=self.queued(db);again=self.queued(db);self.assertEqual(r,again);self.assertEqual(len(db.docs),1)
        c=claim(db,{'expectedRecord':r,'sourceFingerprint':fingerprint(r)});self.assertTrue(c['claimed'])
        self.assertFalse(claim(db,{'expectedRecord':r,'sourceFingerprint':fingerprint(r)})['claimed'])
        self.assertEqual(self.queued(db),c['record'])
    def test_expired_uncertain_attempt_is_never_claimed_again(self):
        db=Store();r=self.queued(db);c=claim(db,{'expectedRecord':r,'sourceFingerprint':fingerprint(r)})
        db.docs[(COLLECTION,key(r['recordId']))]['delivery']['leaseExpiresAt']=time.time()-1
        expired=copy.deepcopy(db.docs[(COLLECTION,key(r['recordId']))]);held=review(db,{'expectedRecord':expired,'reason':'expired'})
        self.assertEqual(held['record']['delivery']['state'],'needs_reconciliation')
        self.assertFalse(claim(db,{'expectedRecord':held['record'],'sourceFingerprint':fingerprint(r)})['claimed'])
        self.assertFalse(finish(db,{'recordId':r['recordId'],'token':c['token'],'acknowledged':True,'canonicalExcerptId':'EXC'})['saved'])
    def test_sent_and_legacy_history_cannot_be_reset_by_another_save(self):
        db=Store();r={**self.record(),'handoffStatus':'sent','poetryPleaseItemId':'EXC'};db.create_raw_document_if_absent(COLLECTION,key(r['recordId']),r)
        self.assertEqual(self.queued(db),r)
        db=Store();r={**self.record(),'handoffStatus':'failed','errorMessage':'timeout'};db.create_raw_document_if_absent(COLLECTION,key(r['recordId']),r)
        held=self.queued(db);self.assertEqual(held['delivery']['state'],'needs_reconciliation');self.assertEqual(held['delivery']['attempts'],1)
    def test_changed_source_updates_only_an_unattempted_version_and_invalidates_active_lease(self):
        db=Store();old=self.queued(db);new={**self.record(),'excerpt':'Corrected approved lines'}
        changed=queue(db,{'handoffs':[new]})['records'][0];self.assertEqual(changed['delivery']['attempts'],0);self.assertEqual(changed['excerpt'],new['excerpt'])
        c=claim(db,{'expectedRecord':changed,'sourceFingerprint':fingerprint(changed)})
        held=queue(db,{'handoffs':[self.record()]})['records'][0];self.assertEqual(held['delivery']['state'],'needs_source_review')
        self.assertFalse(finish(db,{'recordId':old['recordId'],'token':c['token'],'acknowledged':True,'canonicalExcerptId':'EXC'})['saved'])
    def test_completion_requires_matching_active_token_and_preserves_receiver_identity(self):
        db=Store();r=self.queued(db);c=claim(db,{'expectedRecord':r,'sourceFingerprint':fingerprint(r)})
        self.assertFalse(finish(db,{'recordId':r['recordId'],'token':'wrong','acknowledged':True,'canonicalExcerptId':'EXC'})['saved'])
        done=finish(db,{'recordId':r['recordId'],'token':c['token'],'acknowledged':True,'canonicalExcerptId':'EXC'})
        self.assertEqual(done['record']['handoffStatus'],'sent');self.assertEqual(self.queued(db)['poetryPleaseItemId'],'EXC')
        self.assertEqual(self.queued(db)['delivery']['attempts'],1)
    def test_partial_bad_input_never_starts_a_batch_and_source_changes_cannot_claim(self):
        db=Store()
        with self.assertRaises(ValueError):queue(db,{'handoffs':[self.record(),{**self.record(),'recordId':'bad'}]})
        self.assertFalse(db.docs);r=self.queued(db)
        self.assertFalse(claim(db,{'expectedRecord':r,'sourceFingerprint':'changed'})['claimed'])
    def test_fingerprint_matches_javascript_unicode_fixture(self):
        self.assertEqual(fingerprint({**self.record(),"author":"Ｐｏｅｔ","excerpt":"Ａpproved\u00a0lines"}),'6cb5b9ed1c3bba9768c445dd423f75650fc4086f409c9519fbc7515222441ef4')
    def test_ordinary_upsert_cannot_bypass_a_managed_delivery(self):
        from weaver_runtime_sync import sync_excerpt_handoffs
        db=Store();self.queued(db)
        with self.assertRaisesRegex(ValueError,'guarded queue actions'):
            sync_excerpt_handoffs(db,{'handoffs':[self.record()]})
    def test_new_runtime_actions_cannot_fall_back_to_sqlite_in_production(self):
        import os
        from unittest.mock import patch
        from weaver_runtime_sync import validate_production_handoff_backend
        with patch.dict(os.environ,{'K_SERVICE':'weaver','WEAVER_LEDGER_BACKEND':'sqlite'},clear=False):
            for action in ['queue_video_excerpt_deliveries','get_video_excerpt_delivery_queue','claim_video_excerpt_delivery','finish_video_excerpt_delivery','review_video_excerpt_delivery']:
                with self.assertRaises(Exception):validate_production_handoff_backend(action)

if __name__=='__main__':unittest.main()
