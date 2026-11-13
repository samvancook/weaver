import copy
import hashlib
import time
import unittest
from weaver_runtime_sync import (claim_video_pipeline,save_video_pipeline,enqueue_video_pipeline,get_video_pipeline,reconcile_video_pipeline_gate,FIRESTORE_VIDEO_PIPELINE_COLLECTION,FIRESTORE_VIDEO_PIPELINE_CONTROL,FIRESTORE_VIDEO_CURATION_GATES_COLLECTION)

class Store:
    def __init__(self): self.docs={};self.versions={};self.counter=0
    def get_raw_document_versioned(self,collection,key):
        pair=(collection,key)
        return copy.deepcopy(self.docs.get(pair)),self.versions.get(pair)
    def write_raw_document_if_unchanged(self,collection,key,record,version):
        pair=(collection,key)
        if self.versions.get(pair)!=version:return None
        self.counter+=1;self.docs[pair]=copy.deepcopy(record);self.versions[pair]=str(self.counter)
        return record
    def create_raw_document_if_absent(self,collection,key,record):
        if (collection,key) in self.docs:return None
        return self.write_raw_document_if_unchanged(collection,key,record,None)
    def list_raw_documents(self,collection):return [copy.deepcopy(v)for(c,_),v in self.docs.items()if c==collection]

class VideoPipelineTests(unittest.TestCase):
    def state(self):
        key=hashlib.sha256(b'event|source').hexdigest()[:32]
        return {'pipelineId':key,'prioritySetId':'event','sourceFileId':'source','calendarReady':True}
    def test_exclusive_durable_lease_and_stale_worker_rejection(self):
        db=Store();claim=claim_video_pipeline(db,{'trigger':'test'});self.assertTrue(claim['claimed'])
        self.assertFalse(claim_video_pipeline(db,{})['claimed'])
        self.assertFalse(save_video_pipeline(db,{'token':'wrong','states':[self.state()]})['ok'])
        self.assertTrue(save_video_pipeline(db,{'token':claim['token'],'states':[self.state()],'summary':{'checked':1}})['ok'])
        self.assertEqual(len(get_video_pipeline(db,{})['records']),1)
        newer=claim_video_pipeline(db,{});self.assertTrue(newer['claimed']);self.assertNotEqual(newer['token'],claim['token'])
        self.assertFalse(save_video_pipeline(db,{'token':claim['token'],'states':[]})['ok'])
    def test_expired_lease_cannot_write_and_new_events_invalidate_calendar_readiness(self):
        db=Store();claim=claim_video_pipeline(db,{})
        pair=(FIRESTORE_VIDEO_PIPELINE_CONTROL,'runner');db.docs[pair]['leaseExpiresAt']=time.time()-1
        self.assertFalse(save_video_pipeline(db,{'token':claim['token'],'states':[self.state()]})['ok'])
        claim=claim_video_pipeline(db,{})
        self.assertTrue(enqueue_video_pipeline(db,{'prioritySetId':'event','sourceFileId':'source','trigger':'review_saved'})['queued'])
        pair=(FIRESTORE_VIDEO_PIPELINE_COLLECTION,self.state()['pipelineId']);db.docs[pair]['requestedAt']='2999-01-01T00:00:00Z'
        save_video_pipeline(db,{'token':claim['token'],'states':[self.state()]})
        self.assertFalse(db.docs[pair]['calendarReady'])
    def test_duplicate_or_wrong_source_state_batch_never_writes(self):
        for states in [[self.state(),self.state()],[{**self.state(),'pipelineId':'wrong'}]]:
            db=Store();claim=claim_video_pipeline(db,{})
            with self.assertRaises(ValueError):save_video_pipeline(db,{'token':claim['token'],'states':states})
            self.assertEqual(get_video_pipeline(db,{})['records'],[])
    def gate(self):
        return {'prioritySetId':'event','sourceFileId':'source','videoRecordId':'weaver:video:root','decision':'ready_for_poetry_please','sourceEvent':'Event',
                'publishableAssetUrl':'https://drive.google.com/file/d/final/view','reviewCount':5,'selectedExcerptRecordIds':['caption'],'decidedAt':'original-approval',
                'poetryPleaseHandoff':{'status':'needs_reconciliation','sourceRecordId':'weaver:video:root','attemptId':'original-attempt','error':'timeout'}}
    def receipt(self):
        return {'ok':True,'received':True,'canonicalVideoId':'canonical','sourceRecordId':'weaver:video:root','event':{'sourceEvent':'Event'},
                'media':{'reviewOnly':True,'driveLink':'https://drive.google.com/file/d/final/view'},'review':{'reviewCount':5,'selectedExcerptIdCount':1}}
    def test_verified_positive_receipt_reconciles_local_attempt_and_preserves_approval_and_failure_history(self):
        db=Store();g=self.gate();key=hashlib.sha256(b'event|source').hexdigest()[:32]
        db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,g,None)
        result=reconcile_video_pipeline_gate(db,{'expectedGate':g,'receipt':self.receipt()});self.assertTrue(result['saved'])
        saved=result['record'];self.assertEqual(saved['decidedAt'],'original-approval');self.assertEqual(saved['poetryPleaseHandoff']['status'],'sent_to_poetry_please')
        self.assertEqual(saved['poetryPleaseHandoff']['previousHandoff']['error'],'timeout')
        self.assertEqual(saved['poetryPleaseHandoff']['attemptId'],'original-attempt')
    def test_changed_source_asset_event_context_or_gate_cannot_reconcile(self):
        mutations=[lambda r:r.update(received=False),lambda r:r.update(sourceRecordId='wrong'),lambda r:r['event'].update(sourceEvent='wrong'),
                   lambda r:r['media'].update(driveLink='https://drive.google.com/file/d/wrong/view'),lambda r:r['media'].update(reviewOnly=False),
                   lambda r:r['review'].update(reviewCount=4),lambda r:r['review'].update(selectedExcerptIdCount=0)]
        for mutate in mutations:
            db=Store();g=self.gate();key=hashlib.sha256(b'event|source').hexdigest()[:32];db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,g,None)
            receipt=self.receipt();mutate(receipt)
            self.assertFalse(reconcile_video_pipeline_gate(db,{'expectedGate':g,'receipt':receipt})['saved'])
            self.assertEqual(db.docs[(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key)],g)

if __name__=='__main__':unittest.main()
