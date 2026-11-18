import copy
import unittest
from weaver_runtime_sync import get_video_hosted_media_verification, save_video_hosted_media_verification
from test_video_pipeline import Store

class HostedVideoVerificationTests(unittest.TestCase):
    def record(self):
        return {"schemaVersion":"1","sourceRecordId":"weaver:video:source","canonicalVideoId":"VIDEO",
                "bucket":"poetry-please.firebasestorage.app","objectName":"content-library/videos/video/original.mov",
                "generation":"123","byteCount":42,"md5Hash":"A"*22+"==","sha256":"a"*64,
                "validation":"full_bytes_md5","verifiedAt":"2026-11-18T15:00:00Z"}
    def test_immutable_generation_scoped_receipt_survives_repeat_checks(self):
        db=Store();record=self.record()
        self.assertTrue(save_video_hosted_media_verification(db,{"record":record})["saved"])
        self.assertEqual(get_video_hosted_media_verification(db,record)["record"],record)
        later={**record,"verifiedAt":"2026-11-18T16:00:00Z"}
        saved=save_video_hosted_media_verification(db,{"record":later})
        self.assertFalse(saved["saved"]);self.assertEqual(saved["record"],record)
        self.assertIsNone(get_video_hosted_media_verification(db,{**record,"generation":"124"})["record"])
    def test_crossed_sources_conflicting_bytes_and_approval_or_token_fields_are_rejected(self):
        db=Store();record=self.record();save_video_hosted_media_verification(db,{"record":record});before=copy.deepcopy(db.docs)
        for patch in [{"objectName":"content-library/videos/other/a.mov"},{"byteCount":0},
                      {"sha256":"b"*64},{"calendarReady":True},{"videoUrl":"private-token"},{"byteCount":True}]:
            with self.assertRaises(ValueError):save_video_hosted_media_verification(db,{"record":{**record,**patch}})
            self.assertEqual(db.docs,before)

if __name__ == '__main__':unittest.main()
