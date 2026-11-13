import copy
import unittest
from test_video_pipeline import Store
from weaver_runtime_sync import save_video_media_resolution, get_video_media_resolutions, prepare_resolved_video_gate, video_resolution_key, VIDEO_MEDIA_RESOLUTIONS, FIRESTORE_VIDEO_CURATION_GATES_COLLECTION

class ResolutionStore(Store):
    def get_raw_document(self, collection, key):
        return copy.deepcopy(self.docs.get((collection, key)))

class VideoResolutionTests(unittest.TestCase):
    def candidate(self):
        return {'prioritySetId':'event','sourceFileId':'source','sourceMediaType':'OM','sourceVideoUrl':'https://drive.google.com/file/d/source/view','baseScore':8.5,'reviewCount':5,'excerptRecordIds':['collected'],'poemTitle':'Poem','author':'Poet'}
    def resolution(self):
        return {'type':'recording','schemaVersion':'1','prioritySetId':'event','sourceFileId':'source','reviewedSourceFileId':'source','sourceRole':'OM','status':'resolved','pairing':{'status':'verified'},'primaryFileId':'fv','assets':[{'fileId':'fv','mediaType':'FV','url':'https://drive.google.com/file/d/fv/view'}]}
    def saved(self, db):
        return save_video_media_resolution(db, {'record':self.resolution(),'expectedRevision':0,'actor':'admin@example.com','event':'pairing_confirmed'})
    def test_link_save_does_not_touch_release_gate_or_approvals(self):
        db=ResolutionStore();gate={'decision':'hold','selectedExcerptRecordIds':['approved-existing'],'note':'Keep'};key=video_resolution_key('event','source');db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,gate,None)
        self.assertTrue(self.saved(db)['ok']);self.assertEqual(db.get_raw_document(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key),gate)
        self.assertEqual(len(db.docs),2)
    def test_concurrent_asset_save_rejected_and_history_preserved(self):
        db=ResolutionStore();first=self.saved(db);self.assertEqual(first['record']['revision'],1)
        self.assertTrue(save_video_media_resolution(db,{'record':self.resolution(),'expectedRevision':0})['conflict'])
        second=save_video_media_resolution(db,{'record':self.resolution(),'expectedRevision':1,'actor':'a','event':'refresh'})
        self.assertEqual(second['record']['revision'],2);self.assertEqual(len(second['record']['history']),2)
    def test_only_verified_media_and_standing_policy_create_ready_gate(self):
        db=ResolutionStore();self.saved(db);c=self.candidate();r=prepare_resolved_video_gate(db,{'candidate':c,'expectedGate':None});self.assertTrue(r['saved']);g=r['record'];self.assertEqual(g['decision'],'ready_for_poetry_please');self.assertEqual(g['sourceReviewFileId'],'source');self.assertEqual(g['releaseApproval']['kind'],'standing_policy')
        self.assertEqual(g['selectedExcerptRecordIds'],['collected'])
    def test_quality_restrictions_and_missing_media_do_not_make_ready(self):
        for patch in [{'baseScore':8.49},{'reviewCount':4},{'excerptRecordIds':[]},{'publicationRestricted':True}]:
            db=ResolutionStore();self.saved(db);r=prepare_resolved_video_gate(db,{'candidate':{**self.candidate(),**patch},'expectedGate':None});self.assertFalse(r['ok']);self.assertIsNone(db.get_raw_document(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,video_resolution_key('event','source')))
        db=ResolutionStore();self.assertFalse(prepare_resolved_video_gate(db,{'candidate':self.candidate(),'expectedGate':None})['ok'])
    def test_hold_reject_existing_attempt_and_changed_gate_never_overwritten(self):
        for gate in [{'decision':'hold'},{'decision':'reject'},{'decision':'ready_for_poetry_please','poetryPleaseHandoff':{'status':'sending'}}]:
            db=ResolutionStore();self.saved(db);key=video_resolution_key('event','source');db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,gate,None)
            self.assertFalse(prepare_resolved_video_gate(db,{'candidate':self.candidate(),'expectedGate':gate})['ok']);self.assertEqual(db.get_raw_document(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key),gate)
        db=ResolutionStore();self.saved(db);key=video_resolution_key('event','source');gate={'decision':'hold'};db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,gate,None)
        self.assertEqual(prepare_resolved_video_gate(db,{'candidate':self.candidate(),'expectedGate':None})['reason'],'gate_changed')
    def test_verified_fv_resolves_legacy_editing_route_without_erasing_history(self):
        db=ResolutionStore();self.saved(db);key=video_resolution_key('event','source');old={'decision':'send_to_editing','note':'Original','selectedExcerptRecordIds':['chosen'],'history':[{'event':'old'}]};db.write_raw_document_if_unchanged(FIRESTORE_VIDEO_CURATION_GATES_COLLECTION,key,old,None)
        new=prepare_resolved_video_gate(db,{'candidate':self.candidate(),'expectedGate':old})['record'];self.assertEqual(new['note'],'Original');self.assertEqual(new['selectedExcerptRecordIds'],['chosen']);self.assertEqual(new['history'][0],old['history'][0]);self.assertEqual(new['history'][-1]['previousDecision'],'send_to_editing')
    def test_resolution_queue_reads_exact_source_and_event(self):
        db=ResolutionStore();self.saved(db);self.assertEqual(len(get_video_media_resolutions(db,{'prioritySetId':'event','sourceFileId':'source'})['records']),1);self.assertEqual(get_video_media_resolutions(db,{'prioritySetId':'other'})['records'],[])

if __name__=='__main__':unittest.main()
