import copy
import time
import unittest
from weaver_runtime_sync import resolve_video_curation_identity


class GateStore:
    def __init__(self, gate, conflict=False):
        self.gate=copy.deepcopy(gate)
        self.conflict=conflict
        self.writes=0
    def get_raw_document_versioned(self, *_args):
        return copy.deepcopy(self.gate), 'version-one'
    def write_raw_document_if_unchanged(self, _collection, _id, record, version):
        self.writes+=1
        if self.conflict or version!='version-one': return None
        self.gate=copy.deepcopy(record)
        return record


class VideoIdentityReviewTest(unittest.TestCase):
    def setUp(self):
        self.gate={'prioritySetId':'event','sourceFileId':'final','videoRecordId':'weaver:video:final',
                   'decision':'ready_for_poetry_please','decidedBy':'sam','decidedAt':'original-time',
                   'publishableAssetUrl':'https://drive.google.com/file/d/final/view',
                   'selectedExcerptRecordIds':['one','two'],'candidateScore':9.05,'excerptBonus':1,'history':[]}
        self.payload={'prioritySetId':'event','sourceFileId':'final','expectedGate':copy.deepcopy(self.gate),
                      'review':{'sourceReviewFileId':'original','reviewSourceFileIds':['original','replacement'],
                                'reviewedBy':'operator','evidenceUrl':'https://docs.google.com/document/d/evidence/edit',
                                'note':'The reviewed and final recording were compared at three positions and match.'},
                      'checkedSourceFileIds':['final','original','replacement'],
                      'receipts':[{'ok':True,'sourceRecordId':'weaver:video:'+x,'received':False,
                                   'ledger':{'count':0},'checkedAtMs':time.time()*1000}
                                  for x in ['final','original','replacement']]}

    def test_only_identity_and_audit_fields_change(self):
        db=GateStore(self.gate)
        result=resolve_video_curation_identity(db,self.payload)
        self.assertTrue(result['saved'])
        self.assertEqual(db.gate['videoRecordId'],'weaver:video:original')
        self.assertEqual(db.gate['sourceReviewFileId'],'original')
        for key,value in self.gate.items():
            if key not in {'videoRecordId','history'}: self.assertEqual(db.gate[key],value)
        self.assertEqual(db.gate['identityReview']['reviewedBy'],'operator')
        self.assertEqual(db.gate['history'][-1]['event'],'canonical_identity_review')

    def test_changed_gate_and_started_handoff_are_not_overwritten(self):
        for field,value in [('selectedExcerptRecordIds',['changed']),('poetryPleaseHandoff',{'status':'sending'}),
                            ('sourceReviewFileId','another-root'),('decision','hold')]:
            db=GateStore({**self.gate,field:value})
            self.assertFalse(resolve_video_curation_identity(db,self.payload)['saved'])
            self.assertEqual(db.writes,0)

    def test_missing_stale_occupied_and_duplicate_receipts_fail_closed(self):
        for change in ['missing','stale','occupied','attempt','duplicate']:
            p=copy.deepcopy(self.payload)
            if change=='missing': p['receipts'].pop()
            elif change=='stale': p['receipts'][0]['checkedAtMs']-=31000
            elif change=='occupied': p['receipts'][1]['received']=True
            elif change=='attempt': p['receipts'][2]['ledger']['count']=1
            else: p['receipts'][2]=copy.deepcopy(p['receipts'][0])
            db=GateStore(self.gate)
            self.assertFalse(resolve_video_curation_identity(db,p)['saved'])
            self.assertEqual(db.writes,0)

    def test_conditional_write_conflict_is_reported(self):
        db=GateStore(self.gate,True)
        result=resolve_video_curation_identity(db,self.payload)
        self.assertFalse(result['saved'])
        self.assertEqual(result['reason'],'concurrent_identity_change')
        self.assertEqual(db.gate,self.gate)


if __name__=='__main__': unittest.main()
