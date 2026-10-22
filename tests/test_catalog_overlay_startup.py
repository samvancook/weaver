import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import prepare_catalog_overlay


class CatalogOverlayStartupTest(unittest.TestCase):
    def test_one_bad_book_does_not_prevent_service_catalog_startup(self):
        with tempfile.TemporaryDirectory() as directory:
            seed = Path(directory) / "seed.db"
            runtime = Path(directory) / "runtime.db"
            with sqlite3.connect(seed) as connection:
                connection.execute("CREATE TABLE sentinel (value TEXT)")
                connection.execute("INSERT INTO sentinel VALUES ('present')")
            with patch.object(prepare_catalog_overlay, "DEFAULT_DB_PATH", seed), \
                 patch.object(prepare_catalog_overlay, "RUNTIME_DB", runtime), \
                 patch.object(prepare_catalog_overlay, "BOOKS", ("WG",)), \
                 patch.object(prepare_catalog_overlay, "fetch_book", return_value={}), \
                 patch.object(prepare_catalog_overlay, "import_book_payload", side_effect=ValueError("duplicate title")):
                self.assertEqual(prepare_catalog_overlay.prepare(), [])
            with sqlite3.connect(runtime) as connection:
                self.assertEqual(connection.execute("SELECT value FROM sentinel").fetchone()[0], "present")


if __name__ == "__main__":
    unittest.main()
